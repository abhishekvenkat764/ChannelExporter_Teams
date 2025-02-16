from setuptools import setup, find_packages

setup(
    name='ChannelExporter_Teams',  # Replace with your package name
    version='1.6',
    packages=find_packages(),
    include_package_data=True,
    install_requires=[
        'selenium',                # Required for browser automation
        'pandas',                  # For data manipulation
        'beautifulsoup4',         # For HTML parsing
        'psutil',                 # For process management
    ],
    author='Abhishek Venkatachalam',                   # Replace with your name
    author_email='abhishek.venkatachalam06@gmail.com', # Replace with your email
    description='A package to automate Microsoft Teams interactions and collate feedback as df or excel from the desired Channel -> SubChannel.',
    long_description=open('README.md').read(),  # Ensure you have a README.md file
    long_description_content_type='text/markdown',    #url='https://github.com/yourusername/teams_automation',  # Replace with your repository URL
    classifiers=[
        'Programming Language :: Python :: 3',
        'License :: OSI Approved :: MIT License',  # Change as appropriate
        'Operating System :: OS Independent',
    ],
    python_requires='>=3.6',
)
