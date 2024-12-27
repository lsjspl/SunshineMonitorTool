import win32api
import win32con
import pywintypes
from typing import Tuple
import argparse
import sys
import subprocess
import re

def get_primary_monitor():
    """
    获取主显示器的设备名称
    """
    try:
        monitor_info = win32api.EnumDisplayDevices(None, 0)
        if monitor_info.StateFlags & win32con.DISPLAY_DEVICE_PRIMARY_DEVICE:
            return monitor_info.DeviceName
        
        # 如果第一个不是主显示器，继续查找
        i = 1
        while True:
            try:
                monitor_info = win32api.EnumDisplayDevices(None, i)
                if monitor_info.StateFlags & win32con.DISPLAY_DEVICE_PRIMARY_DEVICE:
                    return monitor_info.DeviceName
                i += 1
            except:
                break
        return None
    except Exception as e:
        print(f"获取主显示器信息时发生错误: {str(e)}")
        return None

def change_display_settings(width: int, height: int, refresh_rate: int) -> bool:
    """
    修改主显示器的分辨率和刷新率
    """
    try:
        primary_monitor = get_primary_monitor()
        if not primary_monitor:
            print("未找到主显示器")
            return False
            
        dev_mode = pywintypes.DEVMODEType()
        
        dev_mode.PelsWidth = width
        dev_mode.PelsHeight = height
        dev_mode.DisplayFrequency = refresh_rate
        
        dev_mode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT | win32con.DM_DISPLAYFREQUENCY
        
        result = win32api.ChangeDisplaySettingsEx(primary_monitor, dev_mode, 0)
        
        if result == win32con.DISP_CHANGE_SUCCESSFUL:
            return True
        elif result == win32con.DISP_CHANGE_BADMODE:
            print(f"错误：显示器不支持 {width}x{height} {refresh_rate}Hz 的显示模式")
        elif result == win32con.DISP_CHANGE_BADPARAM:
            print("错误：参数无效")
        elif result == win32con.DISP_CHANGE_BADFLAGS:
            print("错误：标志无效")
        elif result == win32con.DISP_CHANGE_FAILED:
            print("错误：显示驱动程序失败")
        elif result == win32con.DISP_CHANGE_RESTART:
            print("错误：需要重启系统")
        else:
            print(f"错误：未知错误，代码：{result}")
        return False
            
    except Exception as e:
        print(f"修改显示设置时发生错误: {str(e)}")
        if "Access is denied" in str(e):
            print("提示：请以管理员权限运行此程序")
        return False

def get_current_resolution() -> Tuple[int, int, int]:
    """
    获取主显示器的分辨率和刷新率
    """
    try:
        primary_monitor = get_primary_monitor()
        if not primary_monitor:
            print("未找到主显示器")
            return (0, 0, 0)
            
        settings = win32api.EnumDisplaySettings(primary_monitor, win32con.ENUM_CURRENT_SETTINGS)
        return (settings.PelsWidth, settings.PelsHeight, settings.DisplayFrequency)
    except Exception as e:
        print(f"获取当前显示设置时发生错误: {str(e)}")
        return (0, 0, 0)

def get_max_resolution() -> Tuple[int, int, int]:
    """
    获取主显示器支持的最大分辨率和刷新率
    """
    try:
        primary_monitor = get_primary_monitor()
        if not primary_monitor:
            print("未找到主显示器")
            return (0, 0, 0)
            
        max_width = 0
        max_height = 0
        max_refresh = 0
        mode_num = 0
        
        while True:
            try:
                settings = win32api.EnumDisplaySettings(primary_monitor, mode_num)
                if not settings:
                    break
                
                if settings.PelsWidth * settings.PelsHeight > max_width * max_height:
                    max_width = settings.PelsWidth
                    max_height = settings.PelsHeight
                    max_refresh = settings.DisplayFrequency
                elif (settings.PelsWidth * settings.PelsHeight == max_width * max_height and 
                      settings.DisplayFrequency > max_refresh):
                    max_refresh = settings.DisplayFrequency
                
                mode_num += 1
            except:
                break
                
        return (max_width, max_height, max_refresh)
    except Exception as e:
        print(f"获取最大分辨率时发生错误: {str(e)}")
        return (0, 0, 0)

def get_supported_modes(monitor_name: str) -> list:
    """
    获取显示器支持的所有显示模式
    """
    modes = []
    mode_num = 0
    try:
        while True:
            settings = win32api.EnumDisplaySettings(monitor_name, mode_num)
            if not settings:
                break
            modes.append((settings.PelsWidth, settings.PelsHeight, settings.DisplayFrequency))
            mode_num += 1
    except:
        pass
    return modes

def find_closest_resolution(available_modes, target_width, target_height):
    """
    从可用显示模式中���到宽高比最接近，且分辨率小于等于目标值的模式
    
    Args:
        available_modes: 可用显示模式列表，格式为 "宽度x高度"
        target_width: 目标宽度
        target_height: 目标高度
        
    Returns:
        最接近的分辨率字符串，格式为 "宽度x高度"
    """
    closest_mode = None
    min_ratio_diff = float('inf')
    target_ratio = target_width / target_height
    target_pixels = target_width * target_height
    
    # 首先筛选出分辨率小于等于目标值的模式
    smaller_modes = []
    for mode in available_modes:
        width, height = map(int, mode.split('x'))
        if width <= target_width and height <= target_height:
            smaller_modes.append(mode)
    
    # 如果没有找到任何小于等于目标值的模式，使用所有模式
    modes_to_check = smaller_modes if smaller_modes else available_modes
    
    for mode in modes_to_check:
        width, height = map(int, mode.split('x'))
        current_ratio = width / height
        current_pixels = width * height
        
        # 计算宽高比的差异
        ratio_diff = abs(current_ratio - target_ratio)
        
        # 如果宽高比非常接近（差异小于1%）
        if ratio_diff < 0.01:
            # 在宽高比接近的情况下，选择分辨率最接近的
            if closest_mode is None:
                closest_mode = mode
                min_ratio_diff = ratio_diff
            else:
                # 比较分辨率差异
                closest_width, closest_height = map(int, closest_mode.split('x'))
                closest_pixels = closest_width * closest_height
                
                # 选择更接近目标分辨率的模式
                if abs(current_pixels - target_pixels) < abs(closest_pixels - target_pixels):
                    closest_mode = mode
                    min_ratio_diff = ratio_diff
        # 如果还没有找到合适的模式，或者找到了宽高比更接近的模式
        elif closest_mode is None or ratio_diff < min_ratio_diff:
            closest_mode = mode
            min_ratio_diff = ratio_diff
            
    return closest_mode

def process_monitor_settings(args):
    # 获取主显示器
    primary_monitor = get_primary_monitor()
    if not primary_monitor:
        print("错误：未找到主显示器")
        return

    # 列出所有支持的显示模式
    if args.list:
        print("支持的显示模式：")
        modes = get_supported_modes(primary_monitor)
        for width, height, refresh in sorted(modes):
            print(f"{width}x{height} {refresh}Hz")
        return

    # 验证参数组合
    if (args.width or args.height) and not (args.width and args.height):
        print("错误：设置分辨率时必须同时指定宽度(--width)和高度(--height)")
        print("例如：python main.py --width 1920 --height 1080")
        print("可选：添加 --refresh 参数设置刷新率，如：python main.py --width 1920 --height 1080 --refresh 60")
        print("提示：使用 --list 参数查看支持的显示模式")
        return

    # 显示当前设置
    if args.current:
        width, height, refresh = get_current_resolution()
        print(f"当前分辨率: {width}x{height}, 刷新率: {refresh}Hz")
        return

    # 设置为最大分辨率
    if args.max:
        max_width, max_height, max_refresh = get_max_resolution()
        print(f"最大支持分辨率: {max_width}x{max_height}, 刷新率: {max_refresh}Hz")
        if change_display_settings(max_width, max_height, max_refresh):
            print("成功设置为最大分辨率")
        else:
            print("设置失败")
        return

    # 设置为1080p
    if args.fhd:
        if change_display_settings(1920, 1080, 60):
            print("成功设置为1920x1080 60Hz")
        else:
            print("设置失败")
        return

    # 自定义设置
    if args.width and args.height:
        # 获取所有支持的显示模式并排序（从小到大）
        modes = get_supported_modes(primary_monitor)
        # 按分辨率从小到大排序
        modes.sort(key=lambda x: (x[0] * x[1], x[2]))
        
        # 创建唯一的分辨率列表（去除重复的分辨率，只保留不同的宽度x高度组合）
        unique_resolutions = list(set(f"{width}x{height}" for width, height, _ in modes))
        
        # 查找最接近的分辨率
        closest_mode = find_closest_resolution(unique_resolutions, args.width, args.height)
        if closest_mode:
            width, height = map(int, closest_mode.split('x'))
            
            # 找到这个分辨率支持的所有刷新率
            supported_refreshes = [refresh for w, h, refresh in modes if w == width and h == height]
            
            # 如果没有指定刷新率，使用当前的刷新率或最高支持的刷新率
            if not args.refresh:
                _, _, current_refresh = get_current_resolution()
                # 如果当前刷新率不支持，使用最高支持的刷新率
                if current_refresh not in supported_refreshes:
                    args.refresh = max(supported_refreshes)
                else:
                    args.refresh = current_refresh
            
            # 如果目标分辨率与请求的不同，显示提示信息
            if width != args.width or height != args.height:
                print(f"请求的分辨率 {args.width}x{args.height} 不支持")
                print(f"使用最接近的支持分辨率: {width}x{height}")
                print(f"支持的刷新率: {', '.join(map(str, sorted(supported_refreshes)))}Hz")
            
            if change_display_settings(width, height, args.refresh):
                print(f"成功设置为 {width}x{height} {args.refresh}Hz")
            else:
                print("设置失败")
        else:
            print("未找到合适的分辨率模式")
        return

def main():
    # 禁用默认的 help 参数
    parser = argparse.ArgumentParser(description='显示器分辨率和刷新率设置工具', add_help=False)
    
    # 添加自定义的帮助选项
    parser.add_argument('--help', action='help', help='显示帮助信息')
    
    # 其他参数
    parser.add_argument('-c', '--current', action='store_true', help='显示当前分辨率和刷新率')
    parser.add_argument('-m', '--max', action='store_true', help='显示并设置最大分辨率和刷新率')
    parser.add_argument('-w', '--width', type=int, help='设置宽度（像素）')
    parser.add_argument('-h', '--height', type=int, help='设置高度（像素）')
    parser.add_argument('-r', '--refresh', type=int, help='设置刷新率（Hz），可选，默认保持当前刷新率')
    parser.add_argument('-f', '--fhd', action='store_true', help='设置为1920x1080 60Hz')
    parser.add_argument('-l', '--list', action='store_true', help='列出所有支持的显示模式')

    args = parser.parse_args()

    # 如果没有参数，显示帮助信息
    if len(sys.argv) == 1:
        parser.print_help()
        return

    process_monitor_settings(args)

if __name__ == "__main__":
    main()
