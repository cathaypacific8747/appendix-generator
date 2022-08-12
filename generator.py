VERSION = '0.0.1'

import pandas as pd
from rich.console import Console
from difflib import SequenceMatcher as sm
import sys, os, shutil

console = Console(highlight=False)

console.print(f'[light_cyan1 bold]Appendix Generator v{VERSION}[/]', highlight=False)
console.print(f'[grey50]Source Code: https://www.github.com/cathaypacific8747/appendix-generator[/]')
console.print(f'[grey50]License: MIT[/]\n')

outdir = "11 - Appendix"
filepaths = [
    (os.path.join(root, file), file)
    for root, _, files in os.walk('.') for file in files
    if file.endswith('.pdf') and outdir in root
]

def locate(target_filename):
    return [(fp, fn) for fp, fn in filepaths if fn == target_filename]

def suggest(target_filename):
    return '\n'.join(f'- [{distance*100:.0f}% match] {fn} @ {fp}' for fp, fn, distance in sorted([(
        fp, fn, sm(None, target_filename, fn).ratio()
    ) for fp, fn in filepaths], key=lambda x: x[2], reverse=True)[:3])

if __name__ == "__main__":
    if not os.path.isfile("master_record.xlsx"):
        console.print(f'[red]master_record.xlsx not found![/]')
        input("Press any key to exit...")
        sys.exit(1)

    if not os.path.isdir(outdir):
        os.mkdir(outdir)
    
    df = pd.read_excel(
        "master_record.xlsx",
        sheet_name="Appendix Generator",
        skiprows="1:2",
        usecols=[1,2,4,6,8,10,12,14,16,18,20,22,24,26,28,30],
        names=[
            "refnum", "featurenum",
            "wsd_letter", "wsd_drawing",
            "dsd_letter", "dsd_drawing",
            "hd_letter", "hd_drawing",
            "clp_letter", "clp_drawing",
            "hkt_letter", "hkt_drawing",
            "hgc_letter", "hgc_drawing",
            "hkbn_letter", "hkbn_drawing",
        ]
    ).dropna()
    categories = ['wsd', 'dsd', 'hd', 'clp', 'hkt', 'hgc', 'hkbn']
    
    for index, row in df.iterrows():
        refnum, featurenum = row["refnum"], row["featurenum"]
        target_outpath = os.path.join(outdir, f"{refnum}")
        if os.path.isdir(target_outpath) and len(os.listdir(target_outpath)):
            console.print(f'[grey50]{refnum} already processed, skipping...[/]')
            continue
        console.print(f'[bold]#{refnum}: Processing...[/]')
        os.mkdir(target_outpath)

        for category in categories:
            for type in ['letter', 'drawing']:
                for i, target_filename in enumerate(f"{row[f'{category}_{type}']}".split('\n')):
                    prefix = f'[bold]#{refnum} {category} {type}{i+1}[/bold]'
                    if 'NOPLAN' in target_filename or target_filename == '' or target_filename == '0':
                        console.print(f'[grey50]{prefix}: no plan or empty cell, skipping.[/]')
                        continue
                    target_filename += '.pdf'

                    paths = locate(target_filename)
                    if len(paths) == 0:
                        console.print(f'[red]{prefix}: "{target_filename}" not found! Did you mean:\n{suggest(target_filename)}?[/]')
                        continue
                    elif len(paths) > 1:
                        pathlist = "\n".join(f"- {path}" for path in paths)
                        console.print(f'[red]{prefix}: "{target_filename}" found multiple times:\n{pathlist}[/]')
                        continue

                    ofp, ofn = paths[0]
                    nfp = os.path.join(target_outpath, f'{category}_{type}_{ofn}')
                    shutil.copy(ofp, nfp)
                    console.print(f'[green]{prefix}: copied successfully to "{nfp}"![/]')
        
        if len(os.listdir(target_outpath)) == 0:
            os.rmdir(target_outpath)

input("Process completed, press any key to exit...")
sys.exit(0)