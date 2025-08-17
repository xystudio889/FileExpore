app = FileExpore
obj = main.py
file_version = 1.0.0.0
product_version = 1.0.0
file_description = 小工具集合
product_name = FileExpore
$(app): $(obj)
	python -m nuitka --remove-output --enable-plugin=tk-inter --onefile --company-name="xystudio" --product-name="$(product_name)" --file-version="$(file_version)" --product-version="$(product_version)" --file-description="$(file_description)" --msvc=latest $(obj)