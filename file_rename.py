'''find and replace substrings in a directory'''

import os

# current directory:
dir = os.path.dirname(os.path.realpath(__file__))
asset_dir = os.path.join(dir, 'Assets')

def sanitize(s):
            return (str(s).strip()
                    .lower()
                    .replace(" ", "_")
                    .replace("&", "and")
                    .replace(",", "")
                    .replace("-", "_")
                    .replace("/", ""))

# find-replace dictionary:
find_rep = {
    'fanuc_m710' : sanitize('Fanuc M-710iC-45'),
    'fanuc_m10' : sanitize('Fanuc M-10iD-10L'),
    'fanuc_mate' : sanitize('Fanuc LRMate 200iD-7L'),
    'pinchr_lr_and_m10' : 'pinchr'
}

# print(f'***Asset path: {asset_dir}')

for old_sub, clean_sub in find_rep.items():
    print(f'\nWorking on replacing all instances of {old_sub} with {clean_sub}\n')
    for file_nm in os.listdir(asset_dir):
        if old_sub in file_nm:
            old_path = os.path.join(asset_dir, file_nm)
            new_nm = file_nm.replace(old_sub, clean_sub)
            new_path = os.path.join(asset_dir, new_nm)
            os.rename(old_path, new_path)
            print(f'Renamed the path {old_path} to {new_path}\n')
