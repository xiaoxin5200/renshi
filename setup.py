from setuptools import setup

setup(
    name="renshi",                    
    version="0.1.0",                  
    author="Xiao Xin",                
    author_email="xiaoxin5200@example.com",
    py_modules=["gui", "main", "utils", "database"],
    package_dir={"": "."},
    entry_points={
        "console_scripts": [
            "renshi=main:main",
        ],
    },
    install_requires=[],     # 空列表，表示无外部依赖
)
