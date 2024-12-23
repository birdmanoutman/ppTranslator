import requests
import json
from typing import Literal, Optional

class OllamaTranslator:
    # 支持的模型列表
    SUPPORTED_MODELS = {
        "qwen:7b": {
            "zh2en_prompt": "Translate this Chinese text to English. Only return the translation:{text}",
            "en2zh_prompt": "Translate this English text to Chinese. Only return the translation:{text}",
            "zh2en_batch_prompt": """Translate these Chinese texts to English. Return each translation separated by "{separator}" without any prefix or explanation:
{text}""",
            "en2zh_batch_prompt": """Translate these English texts to Chinese. Return each translation separated by "{separator}" without any prefix or explanation:
{text}"""
        },
        "llama3:8b": {
            "zh2en_prompt": """<s>[INST] You are a professional translator. Follow these rules strictly:
1. Translate the Chinese text to English
2. Return ONLY the translation
3. No explanations or notes
4. No prefixes like 'Translation:'
5. Keep original punctuation style

Text to translate:
{text}
[/INST]""",
            "en2zh_prompt": """<s>[INST] You are a professional translator. Follow these rules strictly:
1. Translate the English text to Chinese
2. Return ONLY the translation
3. No explanations or notes
4. No prefixes like '翻译:'
5. Keep original punctuation style

Text to translate:
{text}
[/INST]""",
            "zh2en_batch_prompt": """<s>[INST] You are a professional translator. Follow these rules strictly:
1. Translate these Chinese texts to English
2. Return ONLY the translations, separated by "{separator}"
3. No explanations or notes
4. No prefixes like 'Translation:'
5. Keep original punctuation style

Texts to translate:
{text}
[/INST]""",
            "en2zh_batch_prompt": """<s>[INST] You are a professional translator. Follow these rules strictly:
1. Translate these English texts to Chinese
2. Return ONLY the translations, separated by "{separator}"
3. No explanations or notes
4. No prefixes like '翻译:'
5. Keep original punctuation style

Texts to translate:
{text}
[/INST]"""
        }
    }
    
    def __init__(self, model_name: str = "qwen:7b", host: str = "http://localhost:2342"):
        """初始化翻译器
        Args:
            model_name: 模型名称，支持 "qwen:7b" 或 "llama3:8b"
            host: Ollama服务器地址
        """
        if model_name not in self.SUPPORTED_MODELS:
            raise ValueError(f"不支持的模型: {model_name}。支持的模型有: {list(self.SUPPORTED_MODELS.keys())}")
        
        self.model_name = model_name
        self.host = host.rstrip('/')  # 移除末尾的斜杠
        self.api_url = f"{self.host}/api/generate"
        self.model_config = self.SUPPORTED_MODELS[model_name]
    
    def get_prompt(self, text: str, from_lang: str, to_lang: str, is_batch: bool = False) -> str:
        """根据模型和翻译方向获取对应的提示词"""
        if is_batch:
            if from_lang == "zh" and to_lang == "en":
                template = self.model_config["zh2en_batch_prompt"]
            else:
                template = self.model_config["en2zh_batch_prompt"]
        else:
            if from_lang == "zh" and to_lang == "en":
                template = self.model_config["zh2en_prompt"]
            else:
                template = self.model_config["en2zh_prompt"]
        
        return template.format(text=text, separator="|||" if is_batch else "")
    
    def clean_translation(self, text: str, from_lang="zh", to_lang="en") -> str:
        """清理翻译结果中的多余格式"""
        if not text:
            return text
            
        # 清理常见的格式
        text = text.strip()
        
        # 移除模型特定的格式标记
        if self.model_name == "llama3:8b":
            text = text.replace("<s>", "").replace("</s>", "")
            text = text.replace("[INST]", "").replace("[/INST]", "")
            text = text.replace("Assistant:", "").replace("Human:", "")
        
        # 移除翻译相关的前缀
        prefixes = [
            "translation:", "here's the translation:", "translated text:",
            "翻译:", "译文:", "中文翻译:", "英文翻译:",
            "transliteration:", "explanation:", "note:", "chinese:", "english:"
        ]
        for prefix in prefixes:
            if text.lower().startswith(prefix):
                text = text[len(prefix):].strip()
        
        # 移除多余的引号和括号
        text = text.strip('"').strip("'")
        text = text.strip('(').strip(')')
        text = text.strip('[').strip(']')
        
        # 移除常见的说明性文本
        explanations = [
            "only the translation is provided",
            "only translation returned",
            "direct translation:",
            "translated version:",
            "translation result:",
            "chinese text:",
            "english text:",
            "original text:"
        ]
        for exp in explanations:
            text = text.lower().replace(exp, "").strip()
        
        # 如果是中译英，保留英文和标点
        if from_lang == "zh" and to_lang == "en":
            # 保留英文字符、数字、空格和标点符号
            allowed_chars = set('abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 .,!?;:\'"-()[]{}/')
            # 创建一个转换表，保留允许的字符和标点
            text = ''.join(c for c in text if c in allowed_chars or ord(c) < 0x4e00 or ord(c) > 0x9fff)
        elif from_lang == "en" and to_lang == "zh":
            # 如果是英译中，保留中文和基本标点
            allowed_chars = set('，。！？、（）:;,.!?()- ')
            text = ''.join(c for c in text if c in allowed_chars or '\u4e00' <= c <= '\u9fff')
        
        # 清理多余的空格
        text = ' '.join(text.split())
        
        return text.strip()
    
    def test_connection(self) -> bool:
        """测试服务器连接
        Returns:
            bool: 连接是否成功
        """
        try:
            response = requests.get(f"{self.host}/api/version")
            return response.status_code == 200
        except Exception as e:
            print(f"连接测试失败: {str(e)}")
            return False
    
    def translate(self, text: str, from_lang: str = "zh", to_lang: str = "en") -> str:
        """翻译单个文本"""
        if not text or not text.strip():
            return text
        
        # 获取对应的提示词
        prompt = self.get_prompt(text, from_lang, to_lang)
        
        payload = {
            "model": self.model_name,
            "prompt": prompt,
            "stream": False
        }
        
        try:
            response = requests.post(self.api_url, json=payload)
            response.raise_for_status()
            result = response.json()
            translated_text = result.get("response", "").strip()
            
            # 清理翻译结果
            cleaned_text = self.clean_translation(translated_text, from_lang, to_lang)
            
            return cleaned_text if cleaned_text.strip() else f"[Translation Error for: {text}]"
            
        except Exception as e:
            print(f"翻译出错: {str(e)}")
            return f"[Translation Error for: {text}]"

    def batch_translate(self, texts: list, from_lang: str = "zh", to_lang: str = "en") -> list:
        """批量翻译文本列表"""
        if not texts:
            return []
        
        # 用特殊分隔符组合所有文本
        separator = "|||"
        combined_text = f"{separator}".join(texts)
        
        # 获取对应的批量翻译提示词
        prompt = self.get_prompt(combined_text, from_lang, to_lang, is_batch=True)
        
        payload = {
            "model": self.model_name,
            "prompt": prompt,
            "stream": False
        }
        
        try:
            response = requests.post(self.api_url, json=payload)
            response.raise_for_status()
            result = response.json()
            translated_text = result.get("response", "").strip()
            
            # 分割译文
            translations = [t.strip() for t in translated_text.split(separator) if t.strip()]
            
            # 确保每个翻译都有实际内容
            translations = [t for t in translations if t and not t.isspace()]
            
            # 验证和清理每个翻译结果
            cleaned_translations = []
            for i, trans in enumerate(translations):
                # 清理翻译结果
                cleaned_text = self.clean_translation(trans, from_lang, to_lang)
                
                # 确保有实际内容
                if cleaned_text.strip():
                    cleaned_translations.append(cleaned_text)
                else:
                    print(f"警告：第{i+1}个文本的翻译结果为空")
                    if i < len(texts):
                        cleaned_translations.append(f"[Translation Error for: {texts[i]}]")
            
            # 确保译文数量与原文相同
            if len(cleaned_translations) != len(texts):
                print(f"警告：译文数量({len(cleaned_translations)})与原文数量({len(texts)})不匹配")
                print("正在尝试逐个翻译...")
                
                # 如果批量翻译失败，尝试逐个翻译
                while len(cleaned_translations) < len(texts):
                    idx = len(cleaned_translations)
                    single_translation = self.translate(texts[idx], from_lang, to_lang)
                    cleaned_translations.append(single_translation)
            
            return cleaned_translations
            
        except Exception as e:
            print(f"批量翻译出错: {str(e)}")
            print("正在尝试逐个翻译...")
            
            # 如果批量翻译失败，改用逐个翻译
            return [self.translate(text, from_lang, to_lang) for text in texts]