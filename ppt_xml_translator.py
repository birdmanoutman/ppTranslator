"""
PPT XML 翻译器 (PPTXMLTranslator)

功能描述:
---------
这是一个专门用于翻译 PowerPoint (PPTX) 文件的 Python 工具。主要功能包括：

1. 文件处理
   - 解压和重新打包 PPTX 文件
   - XML 解析和修改
   - 临时文件管理

2. 文本处理
   - 提取 PPT 中的所有文本内容
   - 调用翻译服务进行翻译
   - 保持原始格式和样式
   - 在原文下方插入译文

3. 格式优化
   - 智能字体大小调整
     * 原文和译文使用不同的字号
     * 自动选择合适的标准字号
     * 确保最小字体可读性
   - 自动文本框大小调整
     * 自适应文本内容
     * 保持原始位置和对齐方式
   - 支持复杂布局
     * 处理组合形状
     * 保持相对位置关系

4. 配置选项
   - 可自定义翻译服务
   - 可配置字体大小调整策略
   - 支持多语言翻译

使用方法:
--------
1. 命令行使用:
   python ppt_xml_translator.py --input input.pptx --output translated.pptx --from-lang zh --to-lang en

2. 参数说明:
   --input: 输入的 PPTX 文件路径
   --output: 输出的 PPTX 文件路径（可选）
   --from-lang: 源语言（默认：zh）
   --to-lang: 目标语言（默认：en）
   --model: Ollama 模型名称
   --host: Ollama 服务地址

注意事项:
--------
1. 需要确保有足够的磁盘空间用于临时文件
2. 建议在翻译前备份原始文件
3. 处理大型 PPT 文件时可能需要较长时间
4. 某些特殊格式可能需要手动调整

作者: enzo X cursor
版本: 1.0.0
最后更新: 2024-12-21
"""

import xml.etree.ElementTree as ET
import os
import shutil
from typing import Callable, List, Dict
from ollama_service.translate_service import OllamaTranslator
import zipfile
import tempfile

class PPTXMLTranslator:
    def __init__(self, model_name: str = "llama3:8b", host: str = "http://localhost:11434", debug: bool = False):
        """初始化翻译器"""
        self.translator = OllamaTranslator(model_name=model_name, host=host)
        self.debug = debug
        self.namespaces = {
            'p': "http://schemas.openxmlformats.org/presentationml/2006/main",
            'a': "http://schemas.openxmlformats.org/drawingml/2006/main"
        }
        for prefix, uri in self.namespaces.items():
            ET.register_namespace(prefix, uri)
        # PPT标准字号列表（从大到小）
        self.ppt_font_sizes = [72, 48, 44, 40, 36, 32, 28, 24, 20, 18, 16, 14, 12, 11, 10, 9, 8, 7, 6, 5]
        # 定义最小字体大小限制(磅)
        self.min_font_size = 5
    
    def debug_print(self, *args, **kwargs):
        """调试信息打印"""
        if self.debug:
            print(*args, **kwargs)
    
    def copy_element_style(self, source_elem: ET.Element, target_elem: ET.Element):
        """复制元素的样式属性和子元素"""
        if source_elem is not None:
            # 复制属性
            for key, value in source_elem.attrib.items():
                target_elem.set(key, value)
            # 复制子元素
            for child in source_elem:
                target_elem.append(ET.fromstring(ET.tostring(child)))
    
    def find_element_with_style(self, parent: ET.Element, tag: str) -> tuple[ET.Element, ET.Element]:
        """查找带样式的元素
        Args:
            parent: 父元素
            tag: 标签名（不含命名空间）
        Returns:
            (元素, 样式元素)的元组
        """
        elem = parent.find(f".//a:{tag}", self.namespaces)
        style = None
        if elem is not None:
            style = elem.find(f"a:{tag}Pr", self.namespaces)
        return elem, style
    
    def adjust_element_font_size(self, element: ET.Element, attrs: list[str] = None, is_translation: bool = False):
        """调整元素的字体大小
        Args:
            element: 要调整的元素
            attrs: 要调整的属性列表，默认为['sz']
            is_translation: 是否是译文
        """
        if attrs is None:
            attrs = ['sz']
        
        for attr in attrs:
            if attr in element.attrib:
                try:
                    # 尝试将属性值转换为整数
                    size = int(element.attrib[attr])
                    new_size = self.adjust_font_size(size, is_translation)
                    element.set(attr, str(new_size))
                except ValueError:
                    # 如果转换失败，说明是特殊值（如'quarter'），保持原值
                    print(f"警告：遇到特殊字体大小值：{element.attrib[attr]}，保持原值")
                    continue
    
    def create_element_with_style(self, tag: str, parent: ET.Element = None, style_source: ET.Element = None) -> ET.Element:
        """创建带样式的元素
        Args:
            tag: 标签名（不含命名空间）
            parent: 父元素（可选）
            style_source: 样式来源元素（可选）
        Returns:
            创建的元素
        """
        # 创建元素
        if parent is not None:
            new_elem = ET.SubElement(parent, f"{{{self.namespaces['a']}}}{tag}")
        else:
            new_elem = ET.Element(f"{{{self.namespaces['a']}}}{tag}")
        
        # 如果有样式来源，复制样式
        if style_source is not None:
            style_elem = style_source.find(f"a:{tag}Pr", self.namespaces)
            if style_elem is not None:
                new_style = ET.SubElement(new_elem, f"{{{self.namespaces['a']}}}{tag}Pr")
                self.copy_element_style(style_elem, new_style)
        
        return new_elem
    
    def extract_pptx(self, pptx_path: str) -> str:
        """解压PPTX文件到临时目录"""
        # 创建临时目录
        temp_dir = tempfile.mkdtemp(prefix="pptx_")
        
        try:
            # 解压PPTX文件
            with zipfile.ZipFile(pptx_path, 'r') as zip_ref:
                zip_ref.extractall(temp_dir)
            
            return temp_dir
        except Exception as e:
            if os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            raise Exception(f"解压PPTX文件失败: {str(e)}")
    
    def compress_to_pptx(self, dir_path: str, output_pptx: str):
        """将目录压缩为PPTX文件"""
        try:
            with zipfile.ZipFile(output_pptx, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                # 遍历目录中的所有文件
                for root, _, files in os.walk(dir_path):
                    for file in files:
                        # 获取文件的完整路径
                        file_path = os.path.join(root, file)
                        # 获取相对路径（用于ZIP文件中的路径）
                        arcname = os.path.relpath(file_path, dir_path)
                        # 确保使用正斜杠作为路径分隔符
                        arcname = arcname.replace(os.path.sep, '/')
                        # 添加文件到ZIP
                        zip_ref.write(file_path, arcname)
        except Exception as e:
            if os.path.exists(output_pptx):
                os.remove(output_pptx)
            raise Exception(f"创建PPTX文件失败: {str(e)}")
    
    def prepare_output_dir(self, source_dir: str, target_dir: str):
        """准备输出目录"""
        if os.path.exists(target_dir):
            shutil.rmtree(target_dir)
        shutil.copytree(source_dir, target_dir)
    
    def print_element_tree(self, element: ET.Element, level: int = 0):
        """打印元素树结构，用于调试"""
        indent = "  " * level
        tag = element.tag.split('}')[-1]  # 移除命名空间前缀
        attrs = [f"{k}='{v}'" for k, v in element.attrib.items()]
        attrs_str = " ".join(attrs)
        
        if element.text and element.text.strip():
            print(f"{indent}{tag} {attrs_str}: {element.text.strip()}")
        else:
            print(f"{indent}{tag} {attrs_str}")
        
        for child in element:
            self.print_element_tree(child, level + 1)
    
    def find_text_elements(self, root: ET.Element) -> List[Dict]:
        """查找所有文本元素"""
        text_elements = []
        
        # 查找所有文本框（shape）
        shapes = root.findall(".//p:sp", self.namespaces)
        
        for shape in shapes:
            # 获取文本框中的所有段落
            paragraphs = shape.findall(".//a:p", self.namespaces)
            if not paragraphs:
                continue
            
            # 用于存储同一文本框内的所有文本
            shape_text = ""
            shape_style = None
            first_paragraph = None
            
            for p in paragraphs:
                # 打印段落的XML结构
                print("\n=== 段落XML结构 ===")
                self.print_element_tree(p)
                print("==================\n")
                
                # 获取文本运行块
                runs = p.findall(".//a:r", self.namespaces)
                if not runs:
                    continue
                
                # 收集段落中的所有文本
                paragraph_text = ""
                
                for r in runs:
                    t = r.find(".//a:t", self.namespaces)
                    if t is not None and t.text:
                        paragraph_text += t.text
                        # 获取样式信息（如果还没有获取到）
                        if shape_style is None:
                            rPr = r.find("a:rPr", self.namespaces)
                            if rPr is not None and 'sz' in rPr.attrib:
                                shape_style = self.get_paragraph_style(p)
                                print(f"找到文本: {t.text.strip()}, 字体大小: {self.size_to_point(int(rPr.attrib['sz']))}磅")
                            else:
                                rPr = r.find(".//a:rPr", self.namespaces)
                                if rPr is not None and 'sz' in rPr.attrib:
                                    shape_style = self.get_paragraph_style(p)
                                    print(f"找到文本: {t.text.strip()}, 字体大小: {self.size_to_point(int(rPr.attrib['sz']))}磅")
                                else:
                                    print(f"\n=== 运行块XML结构（未找到字体大小） ===")
                                    self.print_element_tree(r)
                                    print("==================\n")
                
                if paragraph_text.strip():
                    # 保存第一个有文本的段落，用于后续复制样式
                    if first_paragraph is None:
                        first_paragraph = p
                    # 添加段落文本到文本框文本，使用换行符分隔
                    if shape_text:
                        shape_text += "\n"
                    shape_text += paragraph_text.strip()
            
            # 如果文本框中有文本，添加到结果列表
            if shape_text.strip():
                # 如果还没有找到样式信息，从第一个段落获取
                if shape_style is None and first_paragraph is not None:
                    shape_style = self.get_paragraph_style(first_paragraph)
                    if 'font_size' not in shape_style:
                        print(f"警告: 无法找到本 '{shape_text.strip()}' 的字体大小设置")
                
                text_elements.append({
                    'paragraph': first_paragraph,  # 使用第一个段落作为参考
                    'text': shape_text.strip(),
                    'style': shape_style,
                    'shape': shape  # 保存文本框引用
                })
        
        return text_elements
    
    def get_paragraph_style(self, p_element):
        """获取段落的样式信息"""
        # 打印段落的XML结构
        print("\n=== 段落样式查找 ===")
        self.print_element_tree(p_element)
        print("==================")
        
        # 1. 首先从文本运行块rPr中查找
        r_elements = p_element.findall('.//a:r', self.namespaces)
        for r in r_elements:
            r_pr = r.find('a:rPr', self.namespaces)
            if r_pr is not None:
                print("找到文本运行块rPr:")
                self.print_element_tree(r_pr)
                if 'sz' in r_pr.attrib:
                    try:
                        font_size = self.size_to_point(int(r_pr.attrib['sz']))
                        print(f"从文本运行块rPr获取到字体大小: {font_size}磅")
                        return {'font_size': self.point_to_size(font_size)}
                    except ValueError:
                        print(f"警告：无法解析字体大小值：{r_pr.attrib['sz']}，尝试其他来源")
                else:
                    print("文本运行块rPr中没有字体大小设置")

        # 2. 从段落的pPr/defRPr中查找
        p_pr = p_element.find('a:pPr', self.namespaces)
        if p_pr is not None:
            print("找到段落pPr:")
            self.print_element_tree(p_pr)
            def_rpr = p_pr.find('a:defRPr', self.namespaces)
            if def_rpr is not None:
                print("找到defRPr:")
                self.print_element_tree(def_rpr)
                if 'sz' in def_rpr.attrib:
                    try:
                        font_size = self.size_to_point(int(def_rpr.attrib['sz']))
                        print(f"从段落pPr/defRPr获取到字体大小: {font_size}磅")
                        return {'font_size': self.point_to_size(font_size)}
                    except ValueError:
                        print(f"警告：无法解析字体大小值：{def_rpr.attrib['sz']}，尝试其他来源")
                else:
                    print("段落pPr/defRPr中没有字体大小设置")
            else:
                print("没有找到defRPr")
        else:
            print("没有找到段落pPr")

        # 3. 从endParaRPr中查找
        end_para_rpr = p_element.find('a:endParaRPr', self.namespaces)
        if end_para_rpr is not None:
            print("找到endParaRPr:")
            self.print_element_tree(end_para_rpr)
            if 'sz' in end_para_rpr.attrib:
                try:
                    font_size = self.size_to_point(int(end_para_rpr.attrib['sz']))
                    print(f"从endParaRPr获取字体大小: {font_size}磅")
                    return {'font_size': self.point_to_size(font_size)}
                except ValueError:
                    print(f"警告：无法解析字体大小值：{end_para_rpr.attrib['sz']}，使用默认值")
            else:
                print("endParaRPr中没有字体大小设置")
        else:
            print("没有找到endParaRPr")

        # 如果所有尝试都失败，使用默认字体大小（18磅）并减小两个标准字号变为14磅）
        default_size = 18.0
        # 找到14磅的位置
        for i, std_size in enumerate(self.ppt_font_sizes):
            if std_size <= 14.0:
                original_size = std_size
                print(f"使用默认字体大小: 18.0磅 -> {original_size}磅")
                return {'font_size': self.point_to_size(original_size)}
        
        # 如果找不到14磅，使用最小字号
        print(f"使用默认字体大小: 18.0磅 -> {self.min_font_size}磅")
        return {'font_size': self.point_to_size(self.min_font_size)}

    def get_font_size_from_rpr(self, rpr_element):
        """从rPr元素中获取字体大小"""
        if rpr_element is None:
            return None
            
        # 直接从属性中获取
        if 'sz' in rpr_element.attrib:
            try:
                return self.size_to_point(int(rpr_element.attrib['sz']))
            except ValueError:
                print(f"警告：无法解析字体大小值：{rpr_element.attrib['sz']}")
                return None
            
        # 从子元素中获取
        sz_element = rpr_element.find('.//a:sz', self.namespaces)
        if sz_element is not None and 'val' in sz_element.attrib:
            try:
                return self.size_to_point(int(sz_element.attrib['val']))
            except ValueError:
                print(f"警告：无法解析字体大小值：{sz_element.attrib['val']}")
                return None
            
        return None
    
    def point_to_size(self, point_size: float) -> int:
        """将磅值转换为XML中的字号值（1磅 = 100单位）"""
        return int(point_size * 100)
    
    def size_to_point(self, size: int) -> float:
        """将XML中的字号值转换为磅值"""
        return size / 100
    
    def get_next_smaller_size(self, current_size: float) -> float:
        """获取下一个更小的标准字号
        Args:
            current_size: 当前字号（磅）
        Returns:
            下一个更小的标准字号（磅）
        """
        # 将当前大小四舍五入到最接近的0.5
        rounded_size = round(current_size * 2) / 2
        
        # 当前大小在标准字号中的位置
        for i, std_size in enumerate(self.ppt_font_sizes):
            if rounded_size >= std_size:
                # 如果找到第一个小于等于当前大小的标准字号
                if i + 2 < len(self.ppt_font_sizes):
                    # 返回后两个标准字号（相当于减小两次）
                    return self.ppt_font_sizes[i + 2]
                else:
                    # 如果已经接近列表末尾，返回最小值
                    return self.min_font_size
        
        # 如果当前大小小于所有标准字号，返回最小值
        return self.min_font_size
        
    def adjust_font_size(self, size: int, is_translation: bool = False) -> int:
        """调整字体大小
        Args:
            size: 原始字体大小（EMU单位）
            is_translation: 是否是译文
        Returns:
            调整后的字体大小（EMU单位）
        """
        # 转换为磅值
        point_size = self.size_to_point(size)
        
        # 获取下一个更小的标准字号
        new_point_size = self.get_next_smaller_size(point_size)
        
        # 如果是译文，再减小一次字号
        if is_translation:
            new_point_size = self.get_next_smaller_size(new_point_size)
        
        # 确保不小于最小字号
        new_point_size = max(new_point_size, self.min_font_size)
        
        # 转换回EMU单位
        return self.point_to_size(new_point_size)
    
    def create_translated_paragraphs(self, original_p: ET.Element, translated_text: str) -> List[ET.Element]:
        """创建翻译后的段落列表，保持原始样式"""
        paragraphs = []
        
        # 获取原始样式元素
        original_r, original_rPr = self.find_element_with_style(original_p, "r")
        
        # 处理每段文本
        texts = translated_text.split('\n')
        for text in texts:
            # 创建新段落
            new_p = self.create_element_with_style("p", style_source=original_p)
            paragraphs.append(new_p)
            
            # 创建文本运行块
            new_r = self.create_element_with_style("r", parent=new_p, style_source=original_r)
            
            # 调整字体大小
            if original_rPr is not None and 'sz' in original_rPr.attrib:
                new_rPr = new_r.find("a:rPr", self.namespaces)
                if new_rPr is not None:
                    self.adjust_element_font_size(new_rPr, is_translation=True)
            
            # 添加文本
            new_t = ET.SubElement(new_r, f"{{{self.namespaces['a']}}}t")
            new_t.text = text
        
        return paragraphs
    
    def set_auto_fit(self, shape_tree: ET.Element):
        """设置文本框自动调整大小"""
        # 包括处理组合形状和普通形状
        def process_shapes(parent_element):
            # 处理普通形状
            for sp in parent_element.findall(".//p:sp", self.namespaces):
                self._set_shape_auto_fit(sp)
            
            # 处理组合形状
            for grp_sp in parent_element.findall(".//p:grpSp", self.namespaces):
                # 移除组合形状的固定大小限制
                grp_sp_pr = grp_sp.find("p:grpSpPr", self.namespaces)
                if grp_sp_pr is not None:
                    xfrm = grp_sp_pr.find("a:xfrm", self.namespaces)
                    if xfrm is not None:
                        # 保存原始变换信息
                        original_off = (xfrm.get('off', ''), )  # 位置偏移
                        original_ext = (xfrm.get('ext', ''), )  # 范围扩展
                        original_choff = (xfrm.get('chOff', ''), )  # 子元素偏移
                        original_chext = (xfrm.get('chExt', ''), )  # 子元素范围
                        
                        # 移除可能限制大小的属性
                        for attr in ['cx', 'cy']:
                            if attr in xfrm.attrib:
                                del xfrm.attrib[attr]
                        
                        # 确保保持相对位置缩放
                        if original_off[0]:
                            xfrm.set('off', original_off[0])
                        if original_ext[0]:
                            xfrm.set('ext', original_ext[0])
                        if original_choff[0]:
                            xfrm.set('chOff', original_choff[0])
                        if original_chext[0]:
                            xfrm.set('chExt', original_chext[0])
                
                # 递归处理组合形状内的所有形状
                process_shapes(grp_sp)
        
        # 开始处理形状树
        process_shapes(shape_tree)
    
    def _set_shape_auto_fit(self, sp: ET.Element):
        """为单个形状设置自动调整属性"""
        tx_body = sp.find(".//a:txBody", self.namespaces)
        if tx_body is not None:
            body_pr = tx_body.find("a:bodyPr", self.namespaces)
            if body_pr is None:
                body_pr = ET.SubElement(tx_body, f"{{{self.namespaces['a']}}}bodyPr")
            
            # 移除所有现有的自动调整设置
            for auto_fit in body_pr.findall("a:noAutofit", self.namespaces):
                body_pr.remove(auto_fit)
            for auto_fit in body_pr.findall("a:normAutofit", self.namespaces):
                body_pr.remove(auto_fit)
            for auto_fit in body_pr.findall("a:spAutoFit", self.namespaces):
                body_pr.remove(auto_fit)
            
            # 移除可能限制自动调整的属性
            for attr in ['w', 'h']:
                if attr in body_pr.attrib:
                    del body_pr.attrib[attr]
            
            # 设置文本框属性
            body_pr.set('wrap', 'square')  # 启用文本换行
            body_pr.set('rtlCol', '0')     # 左到右排列
            body_pr.set('anchor', 'ctr')   # 垂直居中
            body_pr.set('anchorCtr', '1')  # 保持居中
            
            # 添加形状自动调整
            sp_auto_fit = ET.SubElement(body_pr, f"{{{self.namespaces['a']}}}spAutoFit")
            
            # 检查并调整形状属性
            sp_pr = sp.find(".//p:spPr", self.namespaces)
            if sp_pr is not None:
                # 除固定变换属性
                xfrm = sp_pr.find("a:xfrm", self.namespaces)
                if xfrm is not None:
                    # 保存原始变换信息
                    original_off = xfrm.get('off', '')  # 位置偏移
                    original_ext = xfrm.get('ext', '')  # 范围扩展
                    
                    # 移除固定宽度和高度
                    for attr in ['cx', 'cy']:
                        if attr in xfrm.attrib:
                            del xfrm.attrib[attr]
                    
                    # 确保保持相对位置
                    if original_off:
                        xfrm.set('off', original_off)
                    if original_ext:
                        xfrm.set('ext', original_ext)
    
    def translate_slide(self, slide_path: str, translator_func: Callable[[str], str]):
        """翻译单个幻灯片文件"""
        self.debug_print(f"\n处理幻灯片: {slide_path}")
        
        # 解析XML
        tree = ET.parse(slide_path)
        root = tree.getroot()
        
        # 查找所有可能包含字体大小设置的元素
        size_elements = root.findall(".//*[@sz]", self.namespaces)
        self.debug_print(f"找到 {len(size_elements)} 个包含字体大小设置的元素")
        
        # 调整所有字体大小设置
        for elem in size_elements:
            self.adjust_element_font_size(elem)
        
        # 查找所有文本元素
        text_elements = self.find_text_elements(root)
        
        # 处理每个文本元素
        for elem in text_elements:
            original_text = elem['text']
            self.debug_print(f"\n处理文本: {original_text}")
            
            # 获取文本框中的所有段落并整字体大小
            shape = elem['shape']
            for p in shape.findall(".//a:p", self.namespaces):
                for r in p.findall(".//a:r", self.namespaces):
                    rPr = r.find("a:rPr", self.namespaces)
                    if rPr is not None:
                        self.adjust_element_font_size(rPr, ['sz', 'kern', 'spc', 'baseline'])
            
            # 记录原文的换行位置
            line_breaks = []
            current_pos = 0
            for line in original_text.split('\n'):
                current_pos += len(line)
                line_breaks.append(current_pos)
            
            # 获取翻译
            translated_text = translator_func(original_text)
            if not translated_text:
                continue
            
            self.debug_print(f"翻译: {original_text} -> {translated_text}")
            
            # 根据原文的换行位置对译文进行分段
            translated_lines = []
            last_pos = 0
            total_len = len(translated_text)
            ratio = total_len / len(original_text)
            
            for break_pos in line_breaks:
                # 按照原文换行位置的比例计算译文的换行位置
                translated_break_pos = int(break_pos * ratio)
                # 在最接近的单词边界处换行
                if translated_break_pos < total_len:
                    # 向后查找空格或标点
                    space_pos = translated_text.find(' ', translated_break_pos)
                    punct_pos = -1
                    for punct in '.,:;?!':
                        pos = translated_text.find(punct, translated_break_pos)
                        if pos != -1 and (punct_pos == -1 or pos < punct_pos):
                            punct_pos = pos + 1
                    
                    # 使用最近的分隔位置
                    if space_pos != -1 and punct_pos != -1:
                        break_pos = min(space_pos, punct_pos)
                    elif space_pos != -1:
                        break_pos = space_pos
                    elif punct_pos != -1:
                        break_pos = punct_pos
                    else:
                        break_pos = translated_break_pos
                    
                    # 提取这一行文本
                    line = translated_text[last_pos:break_pos].strip()
                    if line:
                        translated_lines.append(line)
                    last_pos = break_pos
            
            # 添加最后一行
            if last_pos < total_len:
                last_line = translated_text[last_pos:].strip()
                if last_line:
                    translated_lines.append(last_line)
            
            # 将处理后的译文重新组合
            processed_translation = '\n'.join(translated_lines)
            
            # 创建翻译后的段落列表
            translated_paragraphs = self.create_translated_paragraphs(elem['paragraph'], processed_translation)
            
            # 将翻译段落插入到最后一个段落后面
            shape = elem['shape']  # 获取文本框引用
            paragraphs = shape.findall(".//a:p", self.namespaces)
            if paragraphs:
                # 找到最后一个段落的父元素
                last_p = paragraphs[-1]
                parent = None
                for p in root.findall(".//a:p/..", self.namespaces):
                    if last_p in list(p):
                        parent = p
                        break
                
                if parent is not None:
                    # 获取最后一个段落的索引
                    children = list(parent)
                    index = children.index(last_p)
                    # 插入所有翻译段落
                    for i, translated_p in enumerate(translated_paragraphs):
                        parent.insert(index + 1 + i, translated_p)
        
        # 设置文本框自动调整
        self.set_auto_fit(root)
        
        # 保存修改
        tree.write(slide_path, encoding="UTF-8", xml_declaration=True)
    
    def translate_pptx(self, input_dir: str, output_dir: str, from_lang="zh", to_lang="en", progress_callback=None):
        """翻译整个PPT文件夹"""
        # 准备输出目录
        self.prepare_output_dir(input_dir, output_dir)
        
        # 遍历处理所有幻灯片文件
        slides_dir = os.path.join(output_dir, "ppt", "slides")
        if os.path.exists(slides_dir):
            try:
                # 获取所有幻灯片文件
                slides = [f for f in os.listdir(slides_dir) 
                         if f.startswith("slide") and f.endswith(".xml")]
                total_slides = len(slides)
                
                for i, filename in enumerate(slides, 1):
                    slide_path = os.path.join(slides_dir, filename)
                    print(f"正在翻译: {filename}")
                    self.translate_slide(
                        slide_path,
                        lambda text: self.translator.translate(text, from_lang=from_lang, to_lang=to_lang)
                    )
                    if progress_callback:
                        progress_callback(i, total_slides)
            except Exception as e:
                raise Exception(f"翻译幻灯片时出错: {str(e)}")
    
    def translate_pptx_file(self, input_pptx: str, output_pptx: str = None, from_lang="zh", to_lang="en", progress_callback=None):
        """翻译PPTX文件"""
        # 如果未指定输出文件路径，在输入文件旁边创建
        if output_pptx is None:
            base_name = os.path.splitext(input_pptx)[0]
            output_pptx = f"{base_name}_translated.pptx"
        
        temp_dir = None
        temp_output_dir = None
        
        try:
            # 解压PPTX
            print(f"正在解压: {input_pptx}")
            temp_dir = self.extract_pptx(input_pptx)
            temp_output_dir = tempfile.mkdtemp(prefix="pptx_translated_")
            
            # 翻译
            print("正在翻译...")
            self.translate_pptx(temp_dir, temp_output_dir, from_lang, to_lang, progress_callback)
            
            # 压缩为新的PPTX
            print(f"正在生成翻译后的文件: {output_pptx}")
            self.compress_to_pptx(temp_output_dir, output_pptx)
            
            print("翻译完成!")
            return output_pptx
            
        except Exception as e:
            # 如果输出文件已经存在且出错，删除它
            if output_pptx and os.path.exists(output_pptx):
                os.remove(output_pptx)
            raise e
            
        finally:
            # 清理临时目录
            if temp_dir and os.path.exists(temp_dir):
                shutil.rmtree(temp_dir)
            if temp_output_dir and os.path.exists(temp_output_dir):
                shutil.rmtree(temp_output_dir)



def main():
    """主函数示例"""
    import argparse
    
    parser = argparse.ArgumentParser(description='PPT翻译工具')
    parser.add_argument('--input', '-i', required=True, help='输入的PPTX文件路径')
    parser.add_argument('--output', '-o', help='输出的PPTX文件路径（可选，默认在输入文件旁边创建）')
    parser.add_argument('--from-lang', default='zh', choices=['zh', 'en'], help='源语言 (默认: zh)')
    parser.add_argument('--to-lang', default='en', choices=['zh', 'en'], help='目标语言 (默认: en)')
    parser.add_argument('--model', default='llama3:8b', help='Ollama模型名称 (默认: llama3:8b)')
    parser.add_argument('--host', default='http://localhost:2342', help='Ollama服务地址 (默认: http://localhost:2342)')
    
    args = parser.parse_args()
    
    # 创建翻译器实例
    translator = PPTXMLTranslator(model_name=args.model, host=args.host)
    
    try:
        # 执行翻译
        print(f"开始处理文件: {args.input}")
        output_file = translator.translate_pptx_file(
            args.input,
            args.output,
            args.from_lang,
            args.to_lang
        )
        print(f"翻译完成! 输出文件: {output_file}")
    except Exception as e:
        print(f"处理过程中出错: {str(e)}")
        raise

if __name__ == "__main__":
    main()



