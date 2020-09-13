from setuptools import setup, find_packages, Extension

with open('README.md', 'r', encoding='utf-8') as f:
    readme = f.read()

ext_modules = [
    Extension('mapikit.macros',
              sources=['ext/macrosmodule.c']),
    Extension('mapikit.callwrapper',
              sources=['ext/callwrappermodule.c']),
]

setup(
    name='mapikit',
    description='Extended MAPI Made Easier',
    author='Nick Czeczulin',
    long_description=readme,
    long_description_content_type='text/markdown',
    license='MIT',
    url='https://github.com/nczeczulin/mapikit',
    classifiers=[
      'Topic :: Communications :: Email',
      'Development Status :: 2 - Pre-Alpha',
      'Intended Audience :: Developers',
      'Natural Language :: English',
      'License :: OSI Approved :: MIT License',
      'Programming Language :: Python',
      'Programming Language :: Python :: 3',
      'Programming Language :: Python :: 3 :: Only',
      'Programming Language :: Python :: 3.8',
      'Programming Language :: Python :: Implementation :: CPython',
      'Operating System :: Microsoft :: Windows'
    ],
    python_requires='>=3.8',
    packages=find_packages(where='src'),
    package_dir={'': 'src'},
    ext_modules=ext_modules,
    use_scm_version=True,
    setup_requires=['setuptools_scm', 'pytest-runner'],
    install_requires=['pywin32>=228'],
    tests_require=['pytest']
)
