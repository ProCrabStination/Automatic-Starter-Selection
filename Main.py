# Main program entry point
# Robert Perry 11/17/2026 starting program creation

# This will handle the pipeline of a job folder to selection of all components.
# this program will need to be installed alongside dependencies, submodules of other projects.

#import all needed modules
import sys
import os
import argparse
import tkinter as tk

#import used submodules
import GUI
import Submodules


if __name__ == "__main__":
    print("Starting program")

    #step one
    print("Step one, create a UI for finding the job folder.")

    selection = GUI.run_gui("Welcome")  # {"job_folder", "pids_pdf", "turnin_pdf"}, or None if cancelled
    if selection is None:
        print("Cancelled by user.")
        sys.exit(0)
    print("Job folder:", selection["job_folder"])
    print("P&ID PDF:", selection["pids_pdf"])
    print("Turn-in PDF:", selection["turnin_pdf"])


