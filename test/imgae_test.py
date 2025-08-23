import os
img_path = r'C:\Users\25096\PycharmProjects\PythonProject\img\donate.png'
img_name = os.path.basename(img_path)  # 去掉目录，只剩文件名
params_text = f"图片: {img_name}"
print(params_text)