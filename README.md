# mysql-dbalib

`mysql-dbalib` is a local Python study package for bundling permitted course
documents into a virtual environment.

The installed Python module is named `mysql_dbalib`.

This package is not an official MySQL project and is not intended to represent
or replace any third-party library.

## Install from local Git

```powershell
pip install git+file:///D:/Dev/Projects/vshp/mysql-dbalib
```

## Documents

Place permitted study materials in:

```text
src/mysql_dbalib/resources/
```

After installation, those files are copied into the installed package under:

```text
site-packages/mysql_dbalib/resources/
```

## Quick check

```powershell
python -c "import mysql_dbalib; print(mysql_dbalib.__version__)"
```
