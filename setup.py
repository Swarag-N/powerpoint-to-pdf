import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="ppt2pdf", # Replace with your own username
    version="0.0.2",
    author="Swarag Narayanasetty",
    description="Converts PPT to PDF",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/Swarag-N/powerpoint-to-pdf",
    packages=setuptools.find_packages(),
    include_package_data=True,
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: BSD License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
    py_modules=['ppt2pdf'],
    install_requires=[
        'click',
        'comtypes'
    ],
    entry_points={
        'console_scripts':[
            'ppt2pdf=ppt2pdf.scripts.cli:cli'
        ]
    }
)