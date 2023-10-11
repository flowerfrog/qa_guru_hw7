import os
import pathlib
from zipfile import ZipFile, ZIP_DEFLATED
import pytest
from path import tmp_path, resources_path, zip_path


@pytest.fixture(scope='session', autouse=True)
def create_zip():
    if not os.path.exists(tmp_path):
        os.mkdir(tmp_path)
    folder = pathlib.Path(resources_path)
    with ZipFile(zip_path, 'w', ZIP_DEFLATED) as zf:
        for file in folder.iterdir():
            zf.write(file, arcname=file.name)
    yield
    os.remove(zip_path)
    os.rmdir(tmp_path)
