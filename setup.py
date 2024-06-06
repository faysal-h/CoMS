from setuptools import setup, find_packages

setup(
    name="CCMS",
    version="1.0.0",
    packages=find_packages(),
    entry_points={
        'console_scripts': [
            'CCMS=CCMS:main',         
                                            
        ],
    },
    author="Faisal Hanif",
    author_email="fayselhanif@gmail.com",
    description="This project aims to eliminate unnecessary manual file work for analysts working on comparison casework by fetching data and making automated worksheets and folders .",
    # long_description=open('README.md').read(),
    # long_description_content_type='text/markdown',
    url="https://github.com/yourusername/my_project",
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.9',
)
