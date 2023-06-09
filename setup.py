import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name='extupdate',
    version='1.0.0',
    description='EXTUPDATE',
    url='https://github.com/sorzkode/',
    author='Mister Riley',
    author_email='<sorzkode@proton.me>',
    packages=setuptools.find_packages(),
    install_requires=['PySimpleGUI'],
    long_description=long_description,
    long_description_content_type="text/markdown",
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
        'Operating System :: OS Independent',
    ],
    license='MIT',
)
