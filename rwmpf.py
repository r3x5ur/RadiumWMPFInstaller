import os
import winreg
from os.path import basename

from lxml import etree
import requests
import platform
import questionary

CONFIG_URL = 'https://dldir1.qq.com/weixin/Windows/XPlugin/updateConfigWin.xml'
WECHATSETUP = 'https://github.com/tom-snow/wechat-windows-versions/releases/download/v3.9.10.27/WeChatSetup-3.9.10.27.exe' # noqa


def query_wechat_version():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Tencent\WeChat', 0, winreg.KEY_READ) as key:
            return winreg.QueryValueEx(key, 'Version')[0]
    except FileNotFoundError:
        return None


def query_wechat_install_path():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r'Software\Tencent\WeChat', 0, winreg.KEY_READ) as key:
            return winreg.QueryValueEx(key, 'InstallPath')[0]
    except FileNotFoundError:
        return None


def get_wechat_uninstall_path():
    install_path = query_wechat_install_path()
    if install_path:
        return os.path.join(install_path, 'Uninstall.exe')


# 版本号算法
def hex_version2str(_str):
    if _str.startswith('0x'):
        _str = _str[2:]
    ver = list(_str)
    ver[0] = '0'
    ver.reverse()
    r = []
    for _ in range(0, len(ver), 2):
        r.append(str(int(f'{ver.pop()}{ver.pop()}', 16)))
    return '.'.join(r)


def get_command(gray_val):
    return f'/plugin set_grayvalue={gray_val}&set_config_url={CONFIG_URL}&check_update_force'


def parse_update_config_xml():
    resp = requests.get(CONFIG_URL)
    xml = etree.XML(resp.text)
    RadiumWMPFList = []
    seen_version = set()
    for item in xml.xpath('//VersionInfo[@name="RadiumWMPF"]'):
        fb = item.xpath('./@forbidSysRegex')
        version = int(item.xpath('./@version')[0])
        app_client_min_version_num = int(item.xpath('./@appClientVerMin')[0], 16)
        app_client_min_version = hex_version2str(hex(app_client_min_version_num))
        grayMax = int(item.xpath('./@grayMax')[0])
        grayMin = int(item.xpath('./@grayMin')[0])
        if version in seen_version:
            continue
        data_item = {
            'url': item.xpath('./@fullurl')[0],
            'forbid_sys_regex': fb[0] if len(fb) else None,
            'version': version,
            'gray_max': grayMax,
            'gray_min': grayMin,
            'app_client_min_version_num': app_client_min_version_num,
            'app_client_min_version': app_client_min_version,
            'is_stable': grayMax == 10000 and grayMin == 1
        }
        seen_version.add(version)
        RadiumWMPFList.append(data_item)
    return RadiumWMPFList


def get_system_info():
    sys = platform.system()
    if sys != 'Windows':
        raise Exception('Only support Windows')
    return f'{sys} {platform.release()}'


def filter_update_config():
    wx_ver = query_wechat_version()
    system_info = get_system_info()
    stable_rwmf = None
    rwmpf_list = []
    for rwmpf in parse_update_config_xml():
        if rwmpf['forbid_sys_regex'] and rwmpf['forbid_sys_regex'] == system_info:
            continue
        if wx_ver < rwmpf['app_client_min_version_num']:
            continue
        if rwmpf['is_stable']:
            if stable_rwmf:
                if stable_rwmf['app_client_min_version_num'] < rwmpf['app_client_min_version_num']:
                    stable_rwmf = rwmpf
            else:
                stable_rwmf = rwmpf
            continue
        rwmpf_list.append(rwmpf)
    rwmpf_list.insert(0, stable_rwmf)
    return rwmpf_list


def picker_version():
    def picker_mapper(item):
        version = str(item['version']).ljust(5)
        gray = str(item["gray_min"]).ljust(5)
        return questionary.Choice([(
            'yellow bold',
            f'RadiumWMPF {version} 灰度值 {gray}'
        )],
            value=item["gray_min"]
        )

    vers = filter_update_config()
    choices = list(map(picker_mapper, vers))
    ROLLBACK_VAL = -65535
    choices.append(questionary.Choice([('red bold', f'没找到想要的版本？(回退微信客户端版本)')], value=ROLLBACK_VAL))
    result = questionary.select(
        "请选择你想安装的 RadiumWMPF 版本",
        choices=choices,
        show_selected=True,
        use_shortcuts=True,
    ).ask()
    if result == ROLLBACK_VAL:
        return f'curl -L "{WECHATSETUP}" -o {basename(WECHATSETUP)}', 'client'
    if not result: return
    return get_command(result), 'RadiumWMPF'


def clean_rwmpf():
    for _ in range(10):
        os.system("taskkill /f /im mmcrashpad_handler64.exe >nul 2>&1")
        os.system("taskkill /f /im WechatAppEx.exe >nul 2>&1")
    # 删除指定目录及其内容
    RadiumWMPFDIR = os.path.join(os.getenv('APPDATA'), 'Tencent', 'WeChat', 'XPlugin', 'Plugins', 'RadiumWMPF')
    for ver in os.listdir(RadiumWMPFDIR):
        _dir = os.path.join(RadiumWMPFDIR, ver)
        os.system(f'del /s /q /f "{_dir}" >nul 2>&1')
        os.system(f'rd /s /q "{_dir}" >nul 2>&1')


def main():
    URL = 'https://github.com/r3x5ur/RadiumWMPFInstaller'
    questionary.print(f'操作文档: {URL}?tab=readme-ov-file#RadiumWMPFInstaller', '#ff0fff bold')
    vn = query_wechat_version()
    if not vn:
        questionary.print(f'当前系统没有安装微信', 'red')
        return
    ver_str = hex_version2str(hex(vn))
    questionary.print(f'当前微信版本：{ver_str}', '#ff0fff')
    result = picker_version()
    if not result: return
    cmd, typ = result
    if typ == 'RadiumWMPF':
        questionary.print('正在清除已安装的RadiumWMPF', 'yellow')
        clean_rwmpf()
        questionary.print('搜索 `:showcmdwnd` 触发`Config Setting`对话框', '#ff0fff')
        questionary.print('请输入以下指令开始更新', '#ff0fff')
    elif typ == 'client':
        questionary.print('回退版本将会退出微信，并卸载在当前版本微信', 'red bold')
        if not questionary.confirm('确定回退吗？', False).ask():
            questionary.print('已取消回退', 'grey')
            return
        questionary.print('正在卸载微信，请记得勾选保留本地数据', 'red bold')
        os.system('taskkill /f /im WeChat.exe >nul 2>&1')
        os.system('taskkill /f /im XWeChat.exe >nul 2>&1')
        clean_rwmpf()
        uninstall = get_wechat_uninstall_path()
        if uninstall:
            os.system(f'start "" "{uninstall}"')
        questionary.print('请执行以下命令开始下载旧版本微信', '#ff0fff')
    questionary.print('=' * 20 + '>', 'grey')
    questionary.print(cmd, 'red bold')
    questionary.print('<' + '=' * 20, 'grey')


if __name__ == '__main__':
    main()
