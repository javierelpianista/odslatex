from setuptools import setup, find_packages

setup(
        name = 'odslatex',
        version = '0.2',
        author = 'Javier Garcia', 
        long_description = open('README.md', 'r').read(),
        packages = find_packages(),
        entry_points = {
            'console_scripts' : [
                'odslatex = odslatex.odslatex:main'
                ]
            },
        install_requires = [
            'typing',
            'numpy'
            ]
        )
