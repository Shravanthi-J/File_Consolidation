import filecombiner, os
from pathlib import Path

if __name__ == '__main__':
    download_path = Path.home() / "Downloads"
    os.chdir(download_path)
    dest_path = input('Enter Destination Path where you want these Assignment Files to be saved:')
    obj=filecombiner.FileCombining(os.listdir(),dest_path)

