# -*- coding: utf-8 -*-
"""
Word LaTeX公式自动转换器
功能：自动打开Word并尝试执行"全部-专业"公式转换
使用COM接口控制Word应用程序
"""

import os
import sys
import time
import win32com.client
from win32com.client import constants

def open_word_and_convert(file_path):
    """打开Word文档并执行全部公式转换为专业格式"""
    
    print("\n" + "="*70)
    print("🚀 Word LaTeX公式自动转换器")
    print("="*70)
    
    # 检查文件是否存在
    if not os.path.exists(file_path):
        print(f"❌ 错误：文件不存在: {file_path}")
        return False
    
    # 获取绝对路径
    abs_path = os.path.abspath(file_path)
    print(f"\n📂 文档路径: {abs_path}")
    
    word = None
    doc = None
    
    try:
        print("\n⏳ 正在启动Word...")
        # 创建Word应用程序实例
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True  # 显示Word窗口
        
        print("⏳ 正在打开文档...")
        # 打开文档
        doc = word.Documents.Open(abs_path)
        
        print("✅ 文档已打开")
        print("\n🔍 正在扫描公式...")
        
        # 等待文档完全加载
        time.sleep(2)
        
        # 获取文档范围内的所有公式
        # 注意：需要通过 Range 或 StoryRanges 访问 OMaths
        formula_count = 0
        all_omaths = []
        
        # 遍历所有故事范围（包括主文档、页眉页脚等）
        for story in doc.StoryRanges:
            while story:
                # 获取当前故事范围中的公式
                if story.OMaths.Count > 0:
                    for i in range(1, story.OMaths.Count + 1):
                        all_omaths.append(story.OMaths.Item(i))
                        formula_count += 1
                
                # 移动到下一个故事范围
                try:
                    story = story.NextStoryRange
                except:
                    break
        
        print(f"📊 检测到 {formula_count} 个公式对象")
        
        if formula_count == 0:
            print("⚠️  警告：文档中没有检测到公式对象")
            print("💡 提示：请确保已经运行过 main.py 将 $...$ 转换为公式对象")
        else:
            print("\n🔄 正在将公式从线性格式转换为专业格式...")
            print("💡 说明：BuildUp() 会将 UnicodeMath 线性文本构建为二维数学公式")
            
            # 遍历所有公式并转换为专业格式
            converted_count = 0
            failed_count = 0
            already_built_count = 0
            
            for i, omath in enumerate(all_omaths, 1):
                try:
                    # 使用BuildUp()方法将线性格式转换为专业格式
                    # BuildUp()方法会将公式从线性格式(如 "(a)/(b)")构建为二维专业格式
                    # 如果公式已经是专业格式，BuildUp()不会报错，只是不会有变化
                    omath.BuildUp()
                    converted_count += 1
                    
                    # 每10个公式显示一次进度
                    if converted_count % 10 == 0:
                        print(f"   ⏳ 已处理 {converted_count}/{formula_count} 个公式...")
                    
                except Exception as e:
                    # 如果公式已经是专业格式，BuildUp可能会失败，这是正常的
                    error_msg = str(e)
                    if "already" in error_msg.lower() or "已经" in error_msg:
                        already_built_count += 1
                    else:
                        failed_count += 1
                        print(f"   ⚠️  公式 {i} 转换失败: {e}")
            
            print(f"\n✅ 转换完成！")
            print(f"📊 统计：")
            print(f"   • 总计: {formula_count} 个公式")
            print(f"   • 成功转换: {converted_count} 个")
            if already_built_count > 0:
                print(f"   • 已是专业格式: {already_built_count} 个")
            if failed_count > 0:
                print(f"   • 失败: {failed_count} 个")
            
            # 询问是否保存
            print("\n💾 Word文档已更新（未保存）")
            print("💡 提示：")
            print("   • 在Word中按 Ctrl+S 保存文档")
            print("   • 或关闭此脚本后手动保存")
            
            # 可选：自动保存
            save_choice = input("\n是否自动保存文档？(y/n): ").strip().lower()
            if save_choice == 'y':
                print("\n⏳ 正在保存...")
                doc.Save()
                print("✅ 文档已保存")
        
        print("\n" + "="*70)
        print("✅ 操作完成")
        print("="*70)
        print("\n💡 Word将保持打开状态，您可以：")
        print("   • 查看转换结果")
        print("   • 手动调整公式格式")
        print("   • 保存或另存为文档")
        print("\n按Enter键退出脚本...")
        input()
        
        return True
        
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        import traceback
        traceback.print_exc()
        return False
    
    finally:
        # 注意：不关闭Word，让用户自己操作
        # 如果需要自动关闭，取消下面的注释
        # if doc:
        #     doc.Close(SaveChanges=False)
        # if word:
        #     word.Quit()
        pass

def main():
    """主函数"""
    # 设置环境变量确保UTF-8编码
    os.environ['PYTHONIOENCODING'] = 'utf-8'

    # 在Windows上强制设置控制台编码
    if sys.platform == 'win32':
        try:
            import ctypes
            kernel32 = ctypes.windll.kernel32
            # 设置控制台输出代码页为UTF-8
            kernel32.SetConsoleOutputCP(65001)
            kernel32.SetConsoleCP(65001)
        except:
            pass
    
    # 检查是否提供了文件路径
    if len(sys.argv) < 2:
        print("\n使用方法:")
        print("  python auto_convert.py <Word文档路径>")
        print("\n或者拖拽Word文档到此脚本")
        
        file_path = input("\n📂 请输入Word文档路径: ").strip().strip('"')
    else:
        file_path = sys.argv[1]
    
    if not file_path:
        print("❌ 错误：未提供文件路径")
        return
    
    # 检查pywin32是否已安装
    try:
        import win32com.client
    except ImportError:
        print("❌ 错误：缺少必要的依赖 pywin32")
        print("\n正在安装 pywin32...")
        import subprocess
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pywin32'])
            print("✅ pywin32 安装完成，请重新运行此脚本")
        except:
            print("❌ 安装失败，请手动运行: pip install pywin32")
        return
    
    # 执行转换
    open_word_and_convert(file_path)

if __name__ == "__main__":
    main()
