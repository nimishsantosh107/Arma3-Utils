import os
import argparse
from typing import List

try:
    import win32com.client as com
    from pywintypes import com_error
    from tqdm import tqdm
    from bs4 import BeautifulSoup
    from rich.console import Console
    from rich.table import Table
except ImportError:
    print(f"Dependencies not found, installing...")

    def __install_reqs():
        import sys
        import subprocess

        requirements = [ "pypiwin32" ,"tqdm" , "beautifulsoup4", "rich"]
        
        for req in requirements:
            subprocess.check_call([sys.executable, "-m", "pip", "install", req])

    __install_reqs()

    print(f"Finished installing dependencies")
    exit(0)


class Constants:
    ARMA_ROOT = "C:\\Program Files\\Steam\\steamapps\\common\\Arma 3"


def file_path(string):
    '''
    Ensures valid path
    '''
    if os.path.isfile(string):
        return string
    else:
        raise FileNotFoundError(string)

def dir_path(string):
    '''
    Ensures valid path
    '''
    if os.path.isdir(string):
        return string
    else:
        raise NotADirectoryError(string)

def _get_dir_size(dir_path):
    '''
    Uses win32 api to check the directory size
    '''    
    fso = com.Dispatch("Scripting.FileSystemObject")
    folder = fso.GetFolder(dir_path)
    return folder.Size / (1024.0 ** 2) # return in MB

def _get_mod_names(soup) -> List[str]:
    '''
    Checks the parsed html file for list of mod names
    '''
    mod_names = []
    for td in soup.find_all("td"):
        if td.parent['data-type'] == "ModContainer":
            try:
                if td['data-type'] == "DisplayName":
                    mod_names.append(td.text)
            except KeyError: pass
    return mod_names

def _get_mod_sizes(mod_names, arma_root) -> List[int]:
    mod_sizes = []
    for mod_name in tqdm(mod_names):
        try:
            mod_size = _get_dir_size(os.path.join(arma_root, '!Workshop', f'@{mod_name}'))
        except com_error:
            mod_name = mod_name.replace(':', '-')
            mod_name = mod_name.replace('/', '-')
            mod_size = _get_dir_size(os.path.join(arma_root, '!Workshop', f'@{mod_name}'))

        mod_sizes.append(mod_size)
    return mod_sizes

def _print_info(mod_names, mod_sizes):
    console = Console()

    table = Table(show_header=True, header_style="bold green")
    table.add_column("No. ")
    table.add_column("Mod Name")
    table.add_column("Size (in MB)", justify="right")

    for i, (name, size) in enumerate(zip(mod_names, mod_sizes)):
        table.add_row(str(i+1), str(name), str(round(size,3)))

    total_size_mb = round(sum(mod_sizes), 2)

    print()
    console.print(f"[bold green]Total Size:[/bold green]: {str(total_size_mb)} MB ({str(round(total_size_mb / 1024.0, 3))} GB)")
    console.print(table)

def _parse_arguments():
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--html-files', nargs="+", type=file_path, help="Path to active modpack html file(s)", required=True)
    parser.add_argument('-a', '--all-mods-file', type=file_path, help="Path to html file containing all downloaded mods", required=True)
    parser.add_argument('-r', '--arma-root', type=dir_path, default=Constants.ARMA_ROOT, help=f"Path to arma root dir. By default it is: {Constants.ARMA_ROOT}")
    return parser.parse_args()

def main():
    args = _parse_arguments()

    active_mod_names = []
    all_mod_names = []
    unused_mod_names = []

    for html_file in args.html_files:
        soup = None
        with open(html_file) as fp:
            soup = BeautifulSoup(fp, "html.parser")
        active_mod_names.extend(_get_mod_names(soup))
    active_mod_names = list(set(active_mod_names))

    with open(args.all_mods_file) as fp:
        soup = BeautifulSoup(fp, "html.parser")
    all_mod_names = _get_mod_names(soup)

    unused_mod_names = [mod for mod in all_mod_names if mod not in active_mod_names]
    unused_mod_sizes = _get_mod_sizes(unused_mod_names, args.arma_root)

    _print_info(unused_mod_names, unused_mod_sizes)

if __name__ == "__main__":
    main()

