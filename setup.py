setup(
    name="nonlinear-equation-solver",
    version="1.0.0",
    description="program for solving a system of equation that is imported from an Excel sheet",
    url="https://github.com/charlieco/Massey-University-nonlinear-equation-solver/blob/main/README.md",
    author="max corpe",
    author_email="charliecorpe@hotmail.com",
    license="MIT",
    classifiers=[
        "License :: OSI Approved :: MIT License",
        "Programming Language :: Python",
        "Programming Language :: Python :: 2",
        "Programming Language :: Python :: 3",
    ],
    packages=["reader"],
    include_package_data=True,
    install_requires=[
        "sympy", "openpyxl", "CoolProp", "numpy","scipy","fluprodia","PySimpleGUI","joblib"
    ],
    entry_points={"console_scripts": ["realpython=main:main"]},
)