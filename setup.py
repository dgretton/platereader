from setuptools import setup, find_packages

setup(
    name='platereader',
    version='1.0',
    packages=find_packages(),
    license='MIT',
    description='',
    long_description=open('README.md').read(),
    install_requires=['pywin32', 'numpy'],
    url='https://github.com/dgretton/platereader.git',
    author='Dana Gretton',
    author_email='dgretton@mit.edu'
)
