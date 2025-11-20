"""Batch runner for FastHenry and FasterCap.

This script reads a list of geometry folders from an Adress.txt file and
runs the existing ``run_fasthenry.py`` and ``run_fastercap.py`` utilities
for each entry. It expects each listed folder to contain a ``FastSolver``
subdirectory with the two generated solver inputs:

* ``Wire_Sections.inp`` (for FastHenry)
* ``Wire_Sections_FastCap.txt`` (for FasterCap)

The outputs ``Zc.mat`` and ``CapacitanceMatrix.txt`` are checked after each
run and warnings are printed if anything is missing.
"""
import os
import sys

from run_fasthenry import run_fasthenry
from run_fastercap import run_fastercap


def read_address_lines(address_file):
    address_file = os.path.abspath(address_file)
    if not os.path.isfile(address_file):
        raise FileNotFoundError(f"Adress.txt not found: {address_file}")

    with open(address_file, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f.readlines()]

    return [line for line in lines if line]


def process_geometry_folder(base_path):
    geometry_root = os.path.abspath(base_path)
    fastsolver_dir = os.path.join(geometry_root, "FastSolver")

    if not os.path.isdir(fastsolver_dir):
        print(f"Warning: FastSolver folder missing for {geometry_root}")
        return

    fh_input = os.path.join(fastsolver_dir, "Wire_Sections.inp")
    fc_input = os.path.join(fastsolver_dir, "Wire_Sections_FastCap.txt")

    if not os.path.isfile(fh_input):
        print(f"Warning: Wire_Sections.inp missing in {fastsolver_dir}")
    else:
        print(f"Running FastHenry for {fh_input}")
        zc_path = run_fasthenry(fh_input)
        if os.path.isfile(zc_path):
            print(f"FastHenry output found: {zc_path}")
        else:
            print(f"Warning: Zc.mat not found after running FastHenry for {fh_input}")

    if not os.path.isfile(fc_input):
        print(f"Warning: Wire_Sections_FastCap.txt missing in {fastsolver_dir}")
    else:
        print(f"Running FasterCap for {fc_input}")
        cap_path = run_fastercap(fc_input)
        if os.path.isfile(cap_path):
            print(f"FasterCap output found: {cap_path}")
        else:
            print(
                "Warning: CapacitanceMatrix.txt not found after running "
                f"FasterCap for {fc_input}"
            )


if __name__ == "__main__":
    if len(sys.argv) >= 2:
        address_txt = sys.argv[1]
    else:
        address_txt = input("Enter path to Adress.txt: ").strip()

    try:
        geometry_folders = read_address_lines(address_txt)
    except FileNotFoundError as exc:
        print(exc)
        sys.exit(1)

    if not geometry_folders:
        print("No geometry folders listed in Adress.txt.")
        sys.exit(0)

    for folder in geometry_folders:
        print("\n=== Processing geometry folder ===")
        print(folder)
        process_geometry_folder(folder)

    print("\nBatch processing complete.")
