import setuptools

setuptools.setup(
    name='extupdate',
    version='1.0.0',
    description='EXTUPDATE',
    url='https://github.com/sorzkode/',
    author='sorzkode',
    author_email='<sorzkode@proton.me>',
    packages=setuptools.find_packages(),
    install_requires=['PySimpleGUI', 'tkinter'],
    long_description='Quickly update/change Excel file extensions.',
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: MIT',
        'Operating System :: OS Independent',
        ],
)