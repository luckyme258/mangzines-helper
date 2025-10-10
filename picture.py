import os
import fitz  # PyMuPDF

def pdf_to_highres_images(pdf_path, output_folder, dpi=300):
    """将PDF直接转为超高清图片"""
    try:
        if not os.path.exists(pdf_path):
            print(f"❌ PDF文件不存在: {pdf_path}")
            return []
            
        # 创建输出文件夹
        os.makedirs(output_folder, exist_ok=True)
        
        # 打开PDF文件
        pdf_document = fitz.open(pdf_path)
        image_paths = []
        
        # 计算缩放比例
        zoom = dpi / 72  # 72是PDF默认DPI
        mat = fitz.Matrix(zoom, zoom)
        
        print(f"🎯 开始转换PDF为超高清图片...")
        print(f"   PDF文件: {os.path.basename(pdf_path)}")
        print(f"   目标DPI: {dpi}")
        print(f"   总页数: {len(pdf_document)}")
        
        for page_num in range(len(pdf_document)):
            # 获取页面
            page = pdf_document[page_num]
            
            # 渲染页面为高清图片
            pix = page.get_pixmap(matrix=mat, alpha=False)
            
            # 保存图片
            output_path = os.path.join(output_folder, f"page_{page_num+1:03d}.png")
            pix.save(output_path)
            image_paths.append(output_path)
            
            print(f"✅ 页面 {page_num+1:02d}: {pix.width}×{pix.height} 像素")
        
        pdf_document.close()
        
        print(f"\n🎉 转换完成！")
        print(f"📊 统计信息:")
        print(f"   - 输入PDF: {pdf_path}")
        print(f"   - 输出图片: {len(image_paths)} 张")
        print(f"   - 图片DPI: {dpi}")
        print(f"   - 输出位置: {os.path.abspath(output_folder)}")
        
        return image_paths
        
    except Exception as e:
        print(f"❌ 转换失败: {str(e)}")
        return []

def main():
    # 配置参数
    pdf_path = "output.pdf"  # 你的PDF文件路径
    output_folder = "highres_images"
    
    # 超高清配置
    TARGET_DPI = 800  # 可以调整为 600 获得更高清效果
    
    print("🚀 PDF转超高清图片工具")
    print("=" * 50)
    
    # 直接转换
    image_paths = pdf_to_highres_images(pdf_path, output_folder, TARGET_DPI)
    
    if not image_paths:
        print("\n💡 提示: 请确保PDF文件存在且路径正确")

if __name__ == "__main__":
    main()