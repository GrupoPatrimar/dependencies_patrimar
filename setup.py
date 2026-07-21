from os import path
from setuptools import setup, find_packages

here = path.abspath(path.dirname(__file__))

with open(path.join(here, 'requirements.txt'), 'r', encoding='utf-16') as requirements_file:
    # Parse requirements.txt, ignoring any commented-out lines.
    requirements = [line for line in requirements_file.read().splitlines()
                    if not line.startswith('#')]
    #requirements = requirements_file.read()

#print(requirements)

setup(
    name='patrimar_dependencies',
    version='2.21.0',
    packages=find_packages(),
    install_requires=requirements,
    author='Renan Oliveira',
    author_email='renanmgx@hotmail.com',
    description='Layout de dependências para projetos Patrimar',
    long_description=open('README.md', encoding='utf-8').read(),
    long_description_content_type='text/markdown',
    url='https://github.com/RenanMGX/dependencies_patrimar.git',
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',
    ],
)
