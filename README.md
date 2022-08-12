# appendix-generator

Appendices for EI/MM reports are notoriously hard to compile by hand. This program aims to make the process easier by automatically searching for them (in deeply nested directories) and compiling them into a single folder.

## For normal usage:

1. Go to the [releases](https://github.com/cathaypacific8747/appendix-generator/releases) section and download the precompiled exe.
2. Put `master_record.xlsx` and the exe into the same directory.
3. Edit the blue-coloured columns in the `Appendix Generator` sheet of `master_record.xlsx`.
4. Run the exe and fix any errors.

## For development:
```
$ pip3 install pipenv
$ pipenv install
$ pipenv run python3 generator.py
PS> pipenv run pyinstaller generator.spec
```

Executable will be generated under `dist/appendix-generator-vX.X.X.exe`