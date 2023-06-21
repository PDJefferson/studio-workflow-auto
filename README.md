# studio-workflow-auto

## Description

This project consists of handling multiple workflows from studio and users using python and argparse for argument parsing. 

The goal is to store metadata from videos to further process them and create thumbnail at specific timeframes for each video. The metadata is stored in a database and the thumbnails are store in a xlsx file



## Installation

runs on a linux environment and requires python3.0 or higher


## Dependencies

- argparse
- sys
- os
- pymongo
- datetime
- ffmpy
- subprocess
- xlsxwriter


## Usage

There are 3 workflows that can be run from the command line. All of them follow the same format:

`python3 main.py --files <path of each file> --xytech <path to xytech> --output <output format [CSV, DB, XLSS]> --verbose`

If the desired output is to get xls file then you must include  the video to process to match the timeframes available in the database based on the timeframes of the video.

`python3 main.py --files <path of each file> --xytech <path to xytech> --output <output format [CSV, DB, XLSS]> --process <path to video> --verbose`


