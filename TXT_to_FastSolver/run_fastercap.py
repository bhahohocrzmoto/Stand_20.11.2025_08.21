import os
import sys
import time
import win32com.client


def run_fastercap(inp_path, options=None):
    """
    Run FasterCap on a single geometry file (.txt) and save the
    Maxwell capacitance matrix as a text file next to the input.

    inp_path : path to FasterCap geometry file (e.g. .txt, NOT list file)
    options  : extra command-line options, e.g. "-a0.01" for 1% rel. accuracy.
               If None or empty, defaults to "-a0.01".
    """
    inp_path = os.path.abspath(inp_path)
    if not os.path.isfile(inp_path):
        raise FileNotFoundError(f"Input file not found: {inp_path}")

    # Default to Auto with max error 1% if nothing is specified
    if options is None or not options.strip():
        options = "-a0.01"

    # Your file is a geometry file, not a list file, so NO -l flag.
    cmdline = f'"{inp_path}" {options}'.strip()

    print(f"Calling FasterCap with:\n  {cmdline}")

    # Create COM object for FasterCap
    FasterCap = win32com.client.Dispatch("FasterCap.Document")

    # Start simulation (ignore return value)
    FasterCap.Run(cmdline)

    # Wait until it finishes – for FasterCap, IsRunning is a method
    start = time.time()
    max_wait_s = 600  # 10 minutes safety

    while True:
        running = FasterCap.IsRunning()
        if not running:
            break

        if time.time() - start > max_wait_s:
            print("Timeout: FasterCap still running after 10 minutes.")
            break

        time.sleep(0.5)

    # Get the capacitance matrix from Automation
    cap = FasterCap.getCapacitance()

    # Save as plain text next to input file
    out_path = os.path.join(os.path.dirname(inp_path), "CapacitanceMatrix.txt")
    with open(out_path, "w", encoding="utf-8") as f:
        for row in cap:
            f.write(" ".join(f"{complex(val):.6e}" for val in row) + "\n")

    print(f"Capacitance matrix written to:\n  {out_path}")
    return out_path


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage:")
        print("  python run_fastercap.py path\\to\\file.txt [extra FasterCap options]")
        sys.exit(1)

    inp = sys.argv[1]
    extra_opts = " ".join(sys.argv[2:])  # optional overrides

    run_fastercap(inp, extra_opts)
