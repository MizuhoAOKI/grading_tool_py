# -*- coding: utf-8 -*-
import os
import sys
import subprocess # To build C program
from subprocess import PIPE
import inquirer   # To use rich CLI
import readchar

# [TODO] This can be set by Excell worksheet.
DEFAULT_EXTENTION = ".c"

# Select file.extention from __candidate / CLI, default extention = ".csv"
def get_path(target_dir, extention=DEFAULT_EXTENTION, objection=""):
    file_list = [] # list up all csv files in ../ref_path
    for file in os.listdir(target_dir):
        base, ext = os.path.splitext(file)
        if ext == extention:
            # file_list.append(os.path.join(target_dir, file)) # absolute path is sometimes too long
            file_list.append(file)
    if len(file_list) < 1 :
        print(f"No {extention} file in {target_dir}")
        readchar.readchar()
        sys.exit()

    cli_selection = [
        inquirer.List(
            'selected_path',
            message='Which file do you run '+ objection +'?',
            choices=file_list,
            carousel=True,
        )
    ]
    answer = inquirer.prompt(cli_selection)
    selected_path = os.path.join(target_dir, answer['selected_path'])
    print(f"Selected path : {selected_path}")
    return selected_path


# Set query to identify which directory you choose.
def execfile(query, root_dir=".", subdir="提出物の添付ファイル"):
    # Show query
    print(f"Searching dir using query : {query}\n")

    # Get the list of students' directories.
    root_path = os.path.abspath(root_dir)
    filesndirs = os.listdir(root_path)
    dirs = [f for f in filesndirs if os.path.isdir(os.path.join(root_path,f))]
    target_dir_candidates = [l for l in dirs if query in l]
    print(target_dir_candidates)
    if len(target_dir_candidates) < 1:
        print(f"Root directory is {root_path}")
        print(f"Query is {query}")
        print("No such directory was found.")
        print("Try another query.")
        readchar.readchar()
        return False
    elif len(target_dir_candidates) > 1:
        print("Couldn't identify single target directory.")
        print("Try another query.")
        readchar.readchar()
        return False

    # Found single target_dir of the specific student you selected.
    target_dir = os.path.abspath(os.path.join(os.path.join(root_dir,target_dir_candidates[0]), subdir))
    target_fn  = os.path.splitext(os.path.basename(get_path(target_dir)))[0] # get filename without extention.
    target_file = os.path.join(target_dir, target_fn)
    print(target_dir)
    print(target_file)

    print("\n##### Build the Program #####")
    print(f"gcc {target_file}.c -o {target_file}")
    proc_build=subprocess.run(f"gcc {target_file}.c -o {target_file}", shell=True, stdout=PIPE, stderr=PIPE, text=True)
    print(proc_build.stdout)
    # print("\n##### Execution Result #####\n")
    # print(f"{target_file}")
    # proc_exec=subprocess.run(os.path.basename(target_file), shell=True, stdin=PIPE, stdout=PIPE, stderr=PIPE, text=True, cwd=target_dir)
    proc_exec=subprocess.run(f"start cmd.exe /k {os.path.basename(target_file)} ", shell=True, stdin=PIPE, stdout=PIPE, stderr=PIPE, text=True, cwd=target_dir)
    print(proc_exec.stdout)

    print("\n############################\n")
    # print("Press Any Key to Exit")
    # readchar.readchar()

if __name__ == '__main__':
    args = sys.argv
    if len(args) == 2:
        execfile(args[1])
    elif 4 <= len(args):
        execfile(args[1], args[2], args[3])
    else:
        print('Arguments are too short')
        readchar.readchar()