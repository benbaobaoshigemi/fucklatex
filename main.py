"""
Word LaTeX公式渲染器 - CLI工具
自动将Word文档中的 $...$ LaTeX公式转换为Word公式对象
"""

import re
import os
import sys
import subprocess
import argparse
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def check_dependencies():
    """检查并安装必要的依赖"""
    required_packages = ['python-docx']
    missing_packages = []

    print("\n🔍 检查依赖...")
    for package in required_packages:
        try:
            __import__(package.replace('-', '_'))
            print(f"   ✅ {package}")
        except ImportError:
            missing_packages.append(package)
            print(f"   ❌ {package} - 未安装")

    if missing_packages:
        print(f"\n⚠️  发现 {len(missing_packages)} 个缺失的依赖包:")
        for pkg in missing_packages:
            print(f"   • {pkg}")

        choice = input("\n是否自动安装缺失的依赖？(y/n): ").strip().lower()
        if choice == 'y':
            print("\n📦 正在安装依赖...")
            try:
                subprocess.check_call([sys.executable, '-m', 'pip', 'install'] + missing_packages)
                print("✅ 依赖安装完成！")
                return True
            except subprocess.CalledProcessError as e:
                print(f"❌ 依赖安装失败: {e}")
                print("请手动运行: pip install " + ' '.join(missing_packages))
                return False
        else:
            print("❌ 请先安装依赖后重新运行程序")
            return False

    print("✅ 所有依赖已就绪！")
    return True

def check_dollar_signs(file_path):
    """检查$符号是否成对出现"""
    try:
        doc = Document(file_path)
        total_dollar_count = 0

        for para in doc.paragraphs:
            text = para.text
            dollar_count = text.count('$')
            total_dollar_count += dollar_count

        if total_dollar_count % 2 != 0:
            print(f"\n❌ 错误：文档中的$符号数量为奇数 ({total_dollar_count})，无法正确匹配公式对！")
            print("请检查文档中的公式格式，确保每个公式都被一对$符号包围。")
            print("\n💡 提示：")
            print("   • 正确格式：$x^2$")
            print("   • 错误格式：$x^2 或 x^2$")
            return False

        print(f"✅ $符号检查通过 (共 {total_dollar_count} 个，{total_dollar_count//2} 对公式)")
        return True

    except Exception as e:
        print(f"❌ 检查$符号时出错: {e}")
        return False

def get_save_mode(input_path):
    """获取保存模式"""
    print(f"\n💾 请选择保存模式:")
    print(f"   0 - 覆盖原文件 (⚠️ 会替换 {input_path})")
    print(f"   1 - 保存到当前目录 ({os.getcwd()})")
    print(f"   2 - 指定保存路径")

    while True:
        choice = input("\n请选择 (0/1/2): ").strip()

        if choice == '0':
            return input_path  # 覆盖原文件

        elif choice == '1':
            # 保存到当前目录
            base_name = os.path.basename(input_path)
            name, ext = os.path.splitext(base_name)
            output_path = os.path.join(os.getcwd(), f"{name}_processed{ext}")
            return output_path

        elif choice == '2':
            # 自定义路径
            output_path = input("请输入完整的保存路径: ").strip().strip('"')
            if not output_path:
                print("❌ 保存路径不能为空")
                continue

            # 确保有扩展名
            if not output_path.lower().endswith('.docx'):
                output_path += '.docx'

            return output_path

        else:
            print("❌ 无效选择，请输入 0、1 或 2")

def latex_to_unicodemath(latex_str):
    """将LaTeX转换为Word的UnicodeMath格式"""
    replacements = {
        r'\\frac\{([^}]+)\}\{([^}]+)\}': r'(\1)/(\2)',
        r'\\sqrt\{([^}]+)\}': r'√(\1)',
        r'\\sum': '∑',
        r'\\int': '∫',
        r'\\prod': '∏',
        r'\\alpha': 'α',
        r'\\beta': 'β',
        r'\\gamma': 'γ',
        r'\\delta': 'δ',
        r'\\epsilon': 'ε',
        r'\\pi': 'π',
        r'\\theta': 'θ',
        r'\\lambda': 'λ',
        r'\\mu': 'μ',
        r'\\sigma': 'σ',
        r'\\omega': 'ω',
        r'\\Gamma': 'Γ',
        r'\\Delta': 'Δ',
        r'\\Theta': 'Θ',
        r'\\Lambda': 'Λ',
        r'\\Sigma': 'Σ',
        r'\\Omega': 'Ω',
        r'\\pm': '±',
        r'\\times': '×',
        r'\\div': '÷',
        r'\\le': '≤',
        r'\\ge': '≥',
        r'\\ne': '≠',
        r'\\approx': '≈',
        r'\\infty': '∞',
        r'\\rightarrow': '→',
        r'\\leftarrow': '←',
        r'\\Rightarrow': '⇒',
        r'\\Leftarrow': '⇐',
    }

    result = latex_str
    for pattern, replacement in replacements.items():
        result = re.sub(pattern, replacement, result)

    return result

def create_omml_formula(latex_str):
    """创建OMML公式元素"""
    try:
        unicodemath = latex_to_unicodemath(latex_str)
        unicodemath_escaped = unicodemath.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')

        omml = f'''<m:oMath {nsdecls("m", "w")}>
    <m:r>
        <w:rPr>
            <w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math" w:cs="Cambria Math"/>
        </w:rPr>
        <m:t>{unicodemath_escaped}</m:t>
    </m:r>
</m:oMath>'''

        return omml, unicodemath
    except Exception as e:
        return None, None

def process_document(input_path, output_path):
    """处理Word文档"""
    try:
        print(f"\n{'='*70}")
        print(f"� 开始处理文档")
        print(f"{'='*70}")
        print(f"📂 输入: {input_path}")
        print(f"📄 输出: {output_path}")

        doc = Document(input_path)
        processed_count = 0
        failed_count = 0

        print(f"\n🔍 扫描文档...")

        for para_idx, para in enumerate(doc.paragraphs):
            text = para.text
            matches = list(re.finditer(r'\$(.*?)\$', text))

            if matches:
                print(f"📝 段落 {para_idx + 1}: {len(matches)} 个公式")

                # 清空段落
                p_element = para._p
                p_element.clear_content()

                # 重建段落
                last_end = 0
                for match in matches:
                    latex_str = match.group(1)
                    start, end = match.span()

                    # 添加公式前的文本
                    if start > last_end:
                        para.add_run(text[last_end:start])

                    # 创建公式
                    omml, unicodemath = create_omml_formula(latex_str)

                    if omml:
                        try:
                            omml_element = parse_xml(omml)
                            para._p.append(omml_element)
                            processed_count += 1
                        except Exception as e:
                            para.add_run(match.group(0))
                            failed_count += 1
                    else:
                        para.add_run(match.group(0))
                        failed_count += 1

                    last_end = end

                # 添加剩余文本
                if last_end < len(text):
                    para.add_run(text[last_end:])

        # 保存文档
        print(f"\n💾 保存文档...")
        doc.save(output_path)

        # 输出结果
        print(f"\n{'='*70}")
        print(f"✅ 处理完成！")
        print(f"{'='*70}")
        print(f"📊 统计:")
        print(f"   • 成功: {processed_count} 个公式")
        if failed_count > 0:
            print(f"   • 失败: {failed_count} 个公式")
        print(f"📄 输出文件: {output_path}")
        print(f"{'='*70}")

        return True

    except Exception as e:
        print(f"\n❌ 处理失败: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """主函数"""
    parser = argparse.ArgumentParser(
        description='Word LaTeX公式渲染器 - 将Word文档中的 $...$ LaTeX公式转换为Word公式对象',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  python main.py                           # 交互模式
  python main.py document.docx             # 处理指定文档（保存到当前目录）
  python main.py document.docx -o output.docx  # 指定输出文件
  python main.py document.docx --overwrite     # 覆盖原文件

保存模式:
  无指定: 保存到当前目录 (filename_processed.docx)
  -o output.docx: 指定输出文件路径
  --overwrite: 覆盖原文件
        """
    )

    parser.add_argument('input_file', nargs='?', help='输入的Word文档路径 (.docx)')
    parser.add_argument('-o', '--output', help='输出文件路径')
    parser.add_argument('--overwrite', action='store_true', help='覆盖原文件')

    args = parser.parse_args()

    print("="*70)
    print(" Word LaTeX公式渲染器 - CLI工具")
    print("="*70)
    print("功能：自动将Word文档中的 $...$ LaTeX公式转换为Word公式对象")
    print("版本：2.0 - CLI版")
    print("="*70)

    # 1. 检查依赖
    if not check_dependencies():
        return

    # 如果提供了命令行参数，使用参数模式
    if args.input_file:
        file_path = args.input_file

        # 验证输入文件
        if not os.path.exists(file_path):
            print(f"❌ 文件不存在: {file_path}")
            return

        if not file_path.lower().endswith('.docx'):
            print("❌ 只支持 .docx 格式的Word文档")
            return

        # 确定输出路径
        if args.overwrite:
            output_path = file_path
        elif args.output:
            output_path = args.output
            if not output_path.lower().endswith('.docx'):
                output_path += '.docx'
        else:
            # 默认保存到当前目录
            base_name = os.path.basename(file_path)
            name, ext = os.path.splitext(base_name)
            output_path = os.path.join(os.getcwd(), f"{name}_processed{ext}")

    else:
        # 交互模式
        # 2. 获取输入文件
        while True:
            file_path = input("\n📂 请输入Word文档路径: ").strip().strip('"')
            if not file_path:
                print("❌ 文件路径不能为空")
                continue

            if not os.path.exists(file_path):
                print(f"❌ 文件不存在: {file_path}")
                continue

            if not file_path.lower().endswith('.docx'):
                print("❌ 只支持 .docx 格式的Word文档")
                continue

            break

        # 3. 检查$符号
        if not check_dollar_signs(file_path):
            return

        # 4. 选择保存模式
        output_path = get_save_mode(file_path)

    # 3/5. 检查$符号（参数模式也需要检查）
    if not check_dollar_signs(file_path):
        return

    # 6. 确认操作（仅交互模式）
    if not args.input_file:
        print(f"\n🔄 准备执行:")
        print(f"   输入文件: {file_path}")
        print(f"   输出文件: {output_path}")

        if output_path == file_path:
            confirm = input("\n⚠️  将覆盖原文件，确定继续？(y/n): ").strip().lower()
            if confirm != 'y':
                print("❌ 操作已取消")
                return

    # 7. 执行处理
    success = process_document(file_path, output_path)

    if success:
        print(f"\n🎉 处理成功！")
        print(f"💡 提示: 打开 {output_path} 查看结果")
    else:
        print(f"\n❌ 处理失败！")

    # 仅在交互模式下等待退出
    if not args.input_file:
        input("\n按Enter键退出...")

if __name__ == "__main__":
    main()
