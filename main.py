import argparse
import sys
import os
import pymongo
import datetime
import ffmpy
import subprocess
import xlsxwriter

parser = argparse.ArgumentParser()
parser.add_argument(
    "--verbose", help="increase output verbosity", action="store_true")
parser.add_argument("--files", help="list of files",
                    nargs="+", required=False)
parser.add_argument("--xytech", help="xytech file", required=False)
parser.add_argument("--output", help="choose between DB or CSV output",
                    choices=["DB", "CSV", "XLS"], required=True)
parser.add_argument("--process", help="video processing", required=False)
args = parser.parse_args()

# assuming the video is 60 frames per second
frame_per_second = 60

# connect to mongo db server
myClient = pymongo.MongoClient("mongodb://localhost:27017/")

# check if the file exists and if the file is the correct file type
def checkFile(files):
    if files is None and args.verbose:
        print(f"{files} is empty")
        return

    if files is list:
        for file in files:
            if not os.path.exists(file):
                if args.verbose:
                    print(f"File {file} does not exist")
                return
    elif files is str:
        if not os.path.exists(files):
            if args.verbose:
                print(f"File {file} does not exist")
        return


def readFile(file):
    if (file is None or file == "" or not os.path.exists(file)):
        if (args.verbose):
            print(f"{file} not found or missing")
        return

    with open(file, 'r') as f:
        return f.read()


def stringIsNumberNotEmptyAndNotSpace(string):
    return (string.isnumeric() and string != "" and not string.isspace())

# split the string by : and returns a dictionary of
# the item that is before the : as the key and the item after the : as the valu
def parseXytechInfo(string):
    if not string:
        raise ValueError("No string passed")
    parsedInfo = {}
    currentKey = ""

    for line in string.splitlines():
        line = line.split(":")

        # if there is no : in the line, and the current key is not in the parsedInfo, then continue
        if (len(line) == 1 and currentKey not in parsedInfo):
            continue

        # if there is a : and the value is empty, then create an empty list
        if ((len(line) > 1) and not line[1] and line[0] not in parsedInfo):
            currentKey = line[0].strip()
            parsedInfo[line[0].strip()] = []
        # if there is a : and the value is not empty, then create a key and
        # add the value to the parsedInfo
        # and set the current key to empty
        elif len(line) > 1 and line[1] and not currentKey:
            currentKey = ""
            parsedInfo[line[0].strip()] = line[1].strip()
        # if there is a : and the value is not empty, and the current key is not empty,
        # then add the value to the list
        elif currentKey in parsedInfo and not line[0].isspace() and line[0]:
            parsedInfo[currentKey].append(line[0].strip())
    return parsedInfo

# splits the string by space and returns a dictionary of the
# first index as the key and the rest of the indexes as the value
def parseBaselightInfo(string):
    if not string:
        if args.verbose:
            print("The string is empty")
        return

    parsedInfo = {}

    for line in string.splitlines():
        line = line.split(" ")

        if not line:
            continue

        key = "/".join(line[0].split("/")[2:])

        if key and key not in parsedInfo:
            parsedInfo[key] = []

        for i in range(1, len(line)):
            if stringIsNumberNotEmptyAndNotSpace(line[i]):
                parsedInfo[key].append(int(line[i]))
    return parsedInfo


def parseFlameInfo(string):
    if not string:
        if args.verbose:
            print("No string passed")
        return
    parsedInfo = {}
    # for each line,
    for line in string.splitlines():

        # get the first path and then the second path
        firstPath, secondPath = line.split(" ", 1)

        # split the second path by space
        secondPath = secondPath.split(" ")

        if not line:
            continue

        # joining the first path with the second one
        key = "/".join((firstPath + " " + secondPath[0]).split("/")[1:])

        if key and key not in parsedInfo:
            parsedInfo[key] = []

        for i in range(1, len(secondPath)):
            if stringIsNumberNotEmptyAndNotSpace(secondPath[i]):
                parsedInfo[key].append(int(secondPath[i]))
    return parsedInfo

# gets the frame and shows the consecutive numbers as ranges
def framesAsRanges(frameList, range):
    if (frameList == None or len(frameList) == 0):
        print("No frameList passed")
        return frameList
    frames = []
    prevValue = currentRange = frameList[0]
    for frame in frameList[1:]:
        if frame == prevValue + range:
            prevValue = frame
        else:
            frames.append(f"{currentRange}-{prevValue}" if currentRange !=
                          prevValue else str(currentRange))
            prevValue = currentRange = frame
    frames.append(f"{currentRange}-{prevValue}" if currentRange !=
                  prevValue else str(currentRange))
    return frames

# merges the two files by using the path of the xytech file and the baselight file;
# will remove the first / from xytech and replace it by the first / of baselight
# merge will be done by using the path of xytech after removng the first and second /
def mergeFilesForXytechAndBaselightByPath(xytech, otherFile):
    locations = xytech['Location']
    stringBuilder = ""
    for location in locations:
        firstPath = location.split("/")[1]
        # from second / to the end of the string
        pathToMatch = "/".join(location.split("/")[3:])
        if pathToMatch in otherFile:
            frames = framesAsRanges(otherFile[pathToMatch], 1)
            for frame in frames:
                # should separate by comma
                stringBuilder += firstPath + "/" + \
                    "/".join(pathToMatch.split("/")) + "," + frame + "\n"
    return stringBuilder


def mergeFilesForXytechAndFlameByPath(
        xytech, otherFile):
    locations = xytech['Location']
    stringBuilder = ""
    for location in locations:
        firstPath = location.split("/")[1]
        # from second / to the end of the string
        pathToMatch = "/".join(location.split("/")[3:])
        for key, item in otherFile.items():
            sec, filePath = key.split(" ")
            if pathToMatch == filePath:
                frames = framesAsRanges(item, 1)
                for frame in frames:
                    # should separate by comma and add the secondary path
                    stringBuilder += sec + " " + firstPath + "/" + \
                        "/".join(pathToMatch.split("/")) + "," + frame + "\n"
    return stringBuilder

# creates a new row for each note in the xytech file
def createNewRowsPerNote(xytech, keys):
    stringBuilder = ""
    columns = ""
    for key in keys:
        if not isinstance(xytech[key], list):
            columns += xytech[key] + ","

    # for each key in the keys list
    for key in keys:
        # if the value is a list
        if isinstance(xytech[key], list):
            # for each value in the list
            for value in xytech[key]:
                # create a new row
                stringBuilder += columns + value + "\n"
    return stringBuilder

# creates the csv file
def createCSVFile(xytech, files):
    if (xytech == None or files == None):
        print("No data passed")
        sys.exit(2)

    # store xytech keys except location
    xytechKeys = [key for key in xytech.keys() if key != "Location"]

    locationsAndFrames = ""
    dateOfFiles = ""
    # for each file in the files dictionary
    for key, file in files.items():
        machine = key.split("_")[0]
        dateOfFiles = key.split("_")[2].split(".")[0]
        currentFile = ""
        if (machine == "Flame"):
            currentFile = mergeFilesForXytechAndFlameByPath(
                xytech, file)
        elif (machine == "Baselight"):
            currentFile = mergeFilesForXytechAndBaselightByPath(
                xytech, file)
        else:
            if (args.verbose):
                print("Machine not supported")
        # sort the locations and frames by the frame number to fix formatting
        currentFile = sorted(currentFile.splitlines(
        ), key=lambda x: int(x.split(",")[1].split("-")[0]))
        locationsAndFrames = locationsAndFrames + "\n" + "\n".join(currentFile)

    with open("output_" + dateOfFiles + ".csv", "w") as f:
        # write row 2 of the csv file the values of the xytech dictionary
        f.write(createNewRowsPerNote(xytech, xytechKeys) + "\n")
        # write row 4 of the csv file the keys of the baselight dictionary
        f.write(locationsAndFrames)


def printWorkDoneByUser(collection, user):
    for workDoneByUser in collection.find({"userOnFile": user}):
        userOnFile = workDoneByUser["userOnFile"]
        dateOfFile = workDoneByUser["dateOfFile"]
        prettyDate = datetime.datetime.fromisoformat(
            dateOfFile).strftime("%m/%d/%Y")
        location = workDoneByUser["location"]
        frame_range = workDoneByUser["frame_range"]
        print("userOnFile: ", userOnFile)
        print("dateOfFile: ", prettyDate)
        print("location: ", location)
        print("frame_range: ", frame_range, "\n")


def printWorkDoneBeforeDateAndMachine(collection, date, machine):
    workDoneBeforeDateAndMachine = collection.aggregate([
        {
            "$lookup": {
                "from": "employee",
                "localField": "userOnFile",
                "foreignField": "userOnFile",
                "as": "emp"
            }

        },
        {
            "$match": {
                "dateOfFile": {"$lt": date.isoformat()},
                "emp.machine": machine
            }
        },
        {
            "$project": {
                "userOnFile": "$userOnFile",
                "location": "$location",
                "frame_range": "$frame_range",
                "dateOfFile": "$dateOfFile",
                "machine": "$emp.machine"
            }
        }
    ]
    )
    for document in workDoneBeforeDateAndMachine:
        userOnFile = document["userOnFile"]
        dateOfFile = document["dateOfFile"]
        frameRange = document["frame_range"]
        location = document["location"]
        dateOfFile = document["dateOfFile"]
        # pretty the dateOfFile which is in datetime format
        prettyDate = datetime.datetime.fromisoformat(
            dateOfFile).strftime("%m/%d/%Y")
        print("machine:", machine)
        print("userOnFile:", userOnFile)
        print("dateOfFile:", prettyDate)
        print("location:", location)
        print("frame/range:", frameRange, "\n")


def printWorkDoneOnAndDate(collection, personComputer, date):
    workDoneInScriptRunnerOnDate = collection.find(
        {"location": {"$regex": ".*" + personComputer + ".*"}, "dateOfFile": date.isoformat()})
    for document in workDoneInScriptRunnerOnDate:
        dateOfFile = document["dateOfFile"]
        frameRange = document["frame_range"]
        location = document["location"]
        dateOfFile = document["dateOfFile"]
        print("Work Done on:", personComputer)
        # pretty the dateOfFile which is in datetime format
        prettyDate = datetime.datetime.fromisoformat(
            dateOfFile).strftime("%m/%d/%Y")
        print("dateOfFile:", prettyDate)
        print("location:", location)
        print("frame range:", frameRange, "\n")


def printAllUsersByMachineType(collection, machine):
    for getOnlyNameByMachineType in collection.distinct("userOnFile",
                                                        {"machine": machine}):
        print(getOnlyNameByMachineType, "\n")


def storeInMongoDB(xytech, files):
    # creates a database called video files
    videoFiles = myClient["videoFiles"]

    # create a collection called employee
    employeeCollection = videoFiles["employee"]

    # create a collection called frame
    frameCollection = videoFiles["frame"]

    # get the script runner from the host machine
    scriptRunner = ""
    try:
        # if using windows
        scriptRunner = os.getlogin()
    except OSError:
        # otherwise in linux
        scriptRunner = subprocess.check_output("whoami").decode("utf-8").strip()

    # current date
    submittedDate = datetime.datetime.now().isoformat()

    for key, file in files.items():
        machine, userOnFile, dateOfFile = key.split("_")
        
        dateOfFile = datetime.datetime.strptime(
            dateOfFile.split(".")[0], "%Y%m%d").isoformat()

        # insert employee data into the employee collection
        employeeCollection.insert_one(
            {"scriptRunner": scriptRunner,
             "machine": machine,
             "userOnFile": userOnFile,
             "dateOfFile": dateOfFile,
             "submittedDate": submittedDate
             }
        )
        if (machine == "Flame"):
            currentFrameAndLocation = mergeFilesForXytechAndFlameByPath(
                xytech, file)
        elif (machine == "Baselight"):
            currentFrameAndLocation = mergeFilesForXytechAndBaselightByPath(
                xytech, file)
        else:
            if (args.verbose):
                print("Machine not supported")

        # insert work done data into the frame collection
        for line in currentFrameAndLocation.splitlines():
            location, frame = line.split(",")
            frameCollection.insert_one(
                {
                    "userOnFile": userOnFile,
                    "dateOfFile": dateOfFile,
                    "location": location,
                    "frame_range": frame
                }
            )


def timeCodeToFrames(timeCode):
    hours, minutes, seconds, frames = timeCode.split(":")
    return int(hours) * 3600 * frame_per_second + \
        int(minutes) * 60 * frame_per_second + \
        int(seconds) * frame_per_second + \
        int(frames)


def frameToTimeCode(frame):
    frameHour = frame // frame_per_second // 60 // 60
    frameSecond = frame // frame_per_second // 60 % 60
    frameMinute = frame // frame_per_second % 60
    frameff = int(frame % frame_per_second)
    return "{:02d}:{:02d}:{:02d}.{:02d}".format(frameHour,  frameSecond, frameMinute, frameff)


# finds the middle frame of a range of frames
def findMiddleFrameFromRange(frameRange):
    start, end = frameRange.split("-")
    return int(start) + (int(end) - int(start)) // 2


# converts seconds to timecode
def secondsToTimeCode(seconds):
    hours = int(seconds / 3600)
    minutes = int(seconds / 60) % 60
    seconds = int(seconds % 60)
    frames = int((seconds-int(seconds)) * frame_per_second)
    return '{:02d}:{:02d}:{:02d}:{:02d}'.format(hours, minutes, seconds, frames)

# finds all frames whithin a video that is less than or equal to the maxFrame passed in
# and returns a list of objects that contains the frame range, middle frame, and time code
def findAllFramesWithinVideo(maxFrame, collection):
    
    # uses a regex function that splits the start and end by a dash and then compare the end
    # to the maxFrame found in the video pass in the process argument
    findWhenRanges = f"Number(this.frame_range.split('-')[1]) <= {maxFrame} && Number(this.frame_range.split('-')[0]) >= 0"

    result = collection.find({"$or": [
        {"frame_range": {"$regex": r"\d+-\d+"}, "$where": findWhenRanges}
    ]})
    list = []
    for document in result:
        frame = document["frame_range"]
        location = document["location"]
        # if the frame has a dash in it, then it is a range of frames
        # and we need to find the middle frame
        if "-" in frame:
            middleFrame = findMiddleFrameFromRange(frame)
            start, end = frame.split("-")
            startTimeCode, endTimeCode = frameToTimeCode(
                int(start)), frameToTimeCode(int(end))
            object = {
                "location": location,
                "frameRange": frame,
                "middleFrame": middleFrame,
                "timeCodeRange": f"{startTimeCode}-{endTimeCode}",
                "timeCode": frameToTimeCode(middleFrame),
            }
        list.append(object)
    return list


def parsedMAchineFiles():
    # read each file and store it in a dictionary
    files = {}
    if (args.files):
        for file in args.files:
            fileName = os.path.basename(file)
            files[fileName] = readFile(file)

    # parsed files
    parsedFiles = {}

    if (args.files):
        # parse the machine files
        for key, file in files.items():
            if (key.startswith("Baselight")):
                parsedFiles[key] = parseBaselightInfo(file)
            elif (key.startswith("Flame")):
                parsedFiles[key] = parseFlameInfo(file)
    return parsedFiles


def parsedXytechFile():
    # xytech file
    xyTechInfo = readFile(args.xytech)

    xyTechParsedInfo = None
    if (xyTechInfo):
        # parse the xytech file
        xyTechParsedInfo = parseXytechInfo(xyTechInfo)
    return xyTechParsedInfo

if (args.output == "CSV"):
    checkFile(args.xytech)
    checkFile(args.files)
    
    xyTechParsedInfo = parsedXytechFile();
    parsedFiles = parsedMAchineFiles();
    
    if (not xyTechParsedInfo and not parsedFiles):
        if (args.verbose):
            print("No files to parse")
        sys.exit(2)
    
    createCSVFile(xyTechParsedInfo, parsedFiles)
    myClient.close()
elif (args.output == "DB"):
    checkFile(args.xytech)
    checkFile(args.files)

    xyTechParsedInfo = parsedXytechFile();
    parsedFiles = parsedMAchineFiles();
    
    if (not xyTechParsedInfo and not parsedFiles):
        if (args.verbose):
            print("No files to read from")
        sys.exit(2)

    storeInMongoDB(xyTechParsedInfo, parsedFiles)
    # print results

    # creates or gets the database called video files
    videoFiles = myClient["videoFiles"]

    # gets the employee collection from the database
    employeeCollection = videoFiles["employee"]

    # gets the frame collection from the database
    frameCollection = videoFiles["frame"]

    print("1). Work done by TDanza\n")
    printWorkDoneByUser(frameCollection, "TDanza")

    print("2). Work done before 3-25-2023 date on a flame machine\n")
    printWorkDoneBeforeDateAndMachine(
        frameCollection, datetime.datetime(2023, 3, 25), "Flame")

    print("3). Work done on hpsans13 on date 3-26-2023\n")
    printWorkDoneOnAndDate(frameCollection, 'hpsans13',
                           datetime.datetime(2023, 3, 26))
    print("4). Name of all users who worked on a flame machine\n")
    printAllUsersByMachineType(employeeCollection, "Flame")
    # close the connection of the database
    myClient.close()
elif (args.output == "XLS"):
    if (args.process and os.path.exists(args.process)):
        
        # creates or gets the database called video files
        videoFiles = myClient["videoFiles"]

        # gets the frame collection from the database
        frameCollection = videoFiles["frame"]

        # runs ffprobe to get the duration in seconds of the video
        ffprobeOutput = ['ffprobe', '-v', 'error', '-show_entries',
                         'format=duration', '-of', 'default=noprint_wrappers=1:nokey=1', args.process]
        
        # get the fps of the video by running ffprobe
        ffProbeFPS = ['ffprobe', '-v', 'error', '-select_streams', 'v:0', '-show_entries',
                            'stream=r_frame_rate', '-of', 'default=noprint_wrappers=1:nokey=1', args.process]
        ffProbeFPS = subprocess.check_output(ffProbeFPS)
        
        # if the ffProbeFPS is not empty, then get the fps of the video otherwise keep it as 60
        if ffProbeFPS:
            frame_per_second = int(ffProbeFPS.decode("utf-8").split("/")[0])

        # Run ffprobe command and capture output
        ffprobeOutput = subprocess.check_output(ffprobeOutput)

        # Convert duration from seconds to timecode format (hh:mm:ss:ff)
        durationSeconds = float(ffprobeOutput.strip())
        timeCode = secondsToTimeCode(durationSeconds)

        informationToStore = findAllFramesWithinVideo(
            timeCodeToFrames(timeCode), frameCollection)

        # create a directory called snapshots if it does not exist to temporarily store all the thumbnails
        if not os.path.exists("snapshots"):
            subprocess.run(["mkdir", "-p", "snapshots"])

        # create a new workbook
        workbook = xlsxwriter.Workbook('video-information.xls')

        # adds a sheet to the workbook
        sheet = workbook.add_worksheet("Video Information")

        # headers for the sheet
        sheet.write(0, 0, "Location")
        sheet.write(0, 1, "Frame Range")
        sheet.write(0, 2, "Time Code Range")
        sheet.write(0, 3, "Thumbnail")

        for i, info in enumerate(informationToStore):
            location = info["location"]
            frameRange = info["frameRange"]
            middleFrame = info["middleFrame"]
            timeCode = info["timeCode"]
            timeCodeRange = info["timeCodeRange"]
            imagePath = f"./snapshots/{middleFrame}.png"
           
           
            ff = ffmpy.FFmpeg(inputs={args.process: None}, outputs={
                              imagePath: f"-ss {timeCode} -vframes 1 -f image2  -r 60 -s 96x74 -y"})
            ff.run()
            sheet.write(i+1, 0, location)
            sheet.write(i+1, 1, frameRange)
            sheet.write(i+1, 2, timeCodeRange)
            # save the image to the sheet
            sheet.insert_image(i+1, 3, imagePath)

        # close the workbook and saves it
        workbook.close()
        # remove the snapshots directory along with all the images in it after the workbook is created
        subprocess.run(["rm", "-rf", "snapshots"])
        myClient.close()
    else:
        if (args.verbose):
            print("No process file specified or missing")
        myClient.close()
        sys.exit(2)
else:
    if (args.verbose):
        print("Output parameter is empty or not supported or not passed in")
    myClient.close()
    sys.exit(2)