import os
import sys
import PyInstaller.__main__

# 当前目录
current_dir = os.path.abspath(os.path.dirname(__file__))

# 执行PyInstaller打包
PyInstaller.__main__.run([
    'ui.py',                                  # 主程序文件
    '--name=AutoPuling',                      # 应用名称
    '--windowed',                             # 无控制台窗口
    f'--icon={os.path.join(current_dir, "puling.ico")}',  # 图标
    '--add-data=puling.ico;.',                # 添加图标资源文件
    '--noconfirm',                            # 不询问确认
    '--clean',                                # 清理临时文件
    # 添加版本信息，使任务栏图标正确显示
    '--version-file=version_info.txt'         # 版本信息文件
])

print("打包完成！") 