import asyncio
import threading
import queue
import re
import os
import fitz  # PyMuPDF
from pyppeteer import launch
from datetime import datetime

class WeChatPDFConverter:
    def __init__(self, input_file="urls.txt"):
        self.input_file = input_file
        self.generated_pdfs = []
        
        # 创建专用事件循环和任务队列
        self.loop = asyncio.new_event_loop()
        self.task_queue = queue.Queue()
        
        # 启动专用事件循环线程
        threading.Thread(target=self.start_async_loop, daemon=True).start()

    def log_message(self, message):
        """在控制台显示日志消息"""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{timestamp}] {message}")

    def start_async_loop(self):
        """启动专用事件循环线程"""
        asyncio.set_event_loop(self.loop)
        self.loop.run_forever()

    def get_urls(self):
        """从文本文件中获取URL列表"""
        try:
            with open(self.input_file, 'r', encoding='utf-8') as f:
                return [url.strip() for url in f.readlines() if url.strip()]
        except Exception as e:
            self.log_message(f"读取URL文件失败: {str(e)}")
            return []

    def start_conversion(self):
        """开始转换处理"""
        # 首先检查是否有URL输入文件
        if os.path.exists(self.input_file):
            urls = self.get_urls()
            if urls:
                self.log_message("开始处理URL...")
                self.generated_pdfs = []  # 重置PDF列表
                
                for url in urls:
                    self.loop.call_soon_threadsafe(asyncio.create_task, self.process_single_pdf(url))
                
                # 等待所有URL转换完成后再合并PDF
                self.loop.call_soon_threadsafe(asyncio.create_task, self.merge_all_pdfs())
                return
                
        # 如果没有URL文件，直接合并当前目录PDF
        self.loop.call_soon_threadsafe(asyncio.create_task, self.merge_all_pdfs())

    async def merge_all_pdfs(self):
        """合并所有PDF（包括URL转换的和本地已有的）"""
        # 等待所有URL导出的PDF生成完成
        if hasattr(self, 'generated_pdfs'):
            while len(self.generated_pdfs) < len(self.get_urls()):
                await asyncio.sleep(0.5)
        
        # 获取所有PDF文件（URL转换的和本地已有的）
        url_pdfs = [pdf[0] for pdf in self.generated_pdfs] if hasattr(self, 'generated_pdfs') else []
        local_pdfs = [f for f in os.listdir('.') if f.lower().endswith('.pdf') and f not in url_pdfs]
        
        # 合并所有PDF（URL转换的在前）
        all_pdfs = url_pdfs + local_pdfs
        
        if not all_pdfs:
            self.log_message("没有找到任何PDF文件")
            return
            
        self.log_message(f"开始合并{len(all_pdfs)}个PDF文件...")
        
        merged_pdf = fitz.open()
        toc = []
        
        for i, pdf_file in enumerate(all_pdfs):
            try:
                with fitz.open(pdf_file) as pdf:
                    start_page = len(merged_pdf)
                    merged_pdf.insert_pdf(pdf)
                    title = os.path.splitext(pdf_file)[0]
                    toc.append([1, f"{i+1}. {title}", start_page + 1])
                    self.log_message(f"✓ 已添加: {pdf_file}")
            except Exception as e:
                self.log_message(f"处理文件 {pdf_file} 时出错: {str(e)}")
                continue
        
        if len(merged_pdf) == 0:
            self.log_message("没有有效的PDF内容可合并")
            return
            
        merged_pdf.set_toc(toc)
        merged_output = "merged_pdfs.pdf"
        merged_pdf.save(merged_output)
        merged_pdf.close()
        
        self.log_message(f"✓ 合并完成: {merged_output}")
        self.log_message(f"共合并了 {len(all_pdfs)} 个PDF文件")

    async def process_single_pdf(self, url):
        """处理单个PDF生成"""
        try:
            self.log_message(f"处理URL: {url}")
            output_path, title = await self.wechat_to_pdf(url)
            self.generated_pdfs.append((output_path, title))
            self.log_message(f"✓ 已生成: {os.path.basename(output_path)}")
        except Exception as e:
            self.log_message(f"✗ 处理失败: {str(e)}")

    async def wechat_to_pdf(self, url, output_path="output.pdf"):
        """将微信公众号文章转换为PDF文件"""
        # 启动浏览器（使用系统Edge浏览器）
        browser = await launch(
            executablePath='C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
            headless=False,
            handleSIGINT=False,
            handleSIGTERM=False,
            handleSIGHUP=False,
            autoClose=False
        )
        page = await browser.newPage()
        
        try:
            # 访问目标页面
            await page.goto(url, {'waitUntil': 'networkidle2', 'timeout': 60000})
            
            # 模拟滚动加载所有内容
            await page.evaluate('''async () => {
                await new Promise((resolve) => {
                    let totalHeight = 0;
                    const distance = 100;
                    const timer = setInterval(() => {
                        const scrollHeight = document.body.scrollHeight;
                        window.scrollBy(0, distance);
                        totalHeight += distance;
                        
                        if(totalHeight >= scrollHeight){
                            clearInterval(timer);
                            resolve();
                        }
                    }, 100);
                });
            }''')
            
            # 获取文章标题
            title = await page.title()
            safe_title = re.sub(r'[\\/*?:"<>|]', "", title)[:50]  # 移除非法字符并截断
            if output_path == "output.pdf":
                output_path = f"{safe_title}.pdf"
            
            # 生成PDF
            pdf_options = {
                'path': output_path,
                'format': 'A4',
                'printBackground': True,
                'margin': {'top': '30px', 'right': '30px', 'bottom': '30px', 'left': '30px'}
            }
            await page.pdf(pdf_options)
            
            return output_path, title
        except Exception as e:
            self.log_message(f"PDF生成过程中发生错误: {str(e)}")
            raise
        finally:
            try:
                # 确保浏览器进程被关闭
                if not page.isClosed():
                    await page.close()
                await browser.close()
                process = browser.process
                if process and process.poll() is None:
                    process.kill()
            except Exception as e:
                self.log_message(f"浏览器关闭时发生错误: {str(e)}")

if __name__ == "__main__":
    converter = WeChatPDFConverter()
    converter.start_conversion()
    
    # 保持程序运行直到所有任务完成
    try:
        while True:
            pass
    except KeyboardInterrupt:
        print("\n程序终止")
