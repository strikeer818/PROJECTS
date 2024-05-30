import csv
import datetime
from http import client
import pymongo
import argparse
import pandas as pd
import os
import subprocess
import xlsxwriter as excel
from frameioclient import FrameioClient
import requests

matched_producer = ''
matched_operator = ''
matched_job = ''
matched_notes = ''
fps = 24

myclient = pymongo.MongoClient("mongodb://localhost:27017/")
mydb = myclient["Project3database"]
mycol1 = mydb["Baselight_export"]
mycol2 = mydb["Xytech"]

parser = argparse.ArgumentParser()
parser.add_argument('--insertBE', action='store_true', dest='INSERTBE', help='insert Baselight_export to database(FOLDER/FRAMES)')
parser.add_argument('--insertX', action='store_true', dest='INSERTX', help='insert Xytech to database(WORKORDER/LOCATION)')
parser.add_argument('--export', action='store_true', dest='EXPORT', help='export to csv file(has the frames csv file)')
parser.add_argument('--process', type=str, dest='PROCESS', help='process video file')
parser.add_argument('--timecode', action='store_true', dest='TIMECODE', help='insert timecode into Baselight_export')
parser.add_argument('--printBE', action='store_true', dest='PRINTBE', help='print data from Baselight_export collection')
parser.add_argument('--printX', action='store_true', dest='PRINTX', help='print data from Xytech collection')
parser.add_argument('--frames', action='store_true', dest='FRAMES', help='Baselight_export frames converted')
parser.add_argument('--output', action='store_true', dest='XLS', help='Export to XLS file similar to proj1 but with timecodes(has the timecodes in the xlsx file)')
parser.add_argument('--exportEX', action='store_true', dest='EXFILE', help='Export Baselight Exports contents')
parser.add_argument('--exportCSV', action='store_true', dest='CSVFILE', help='Export video file contents')
parser.add_argument('--frameio', action='store_true', dest='FRAMEIO', help='export to Frame.io')
parser.add_argument('--processDB', type=str, dest='PROCESSDB', help='process video file contents to database')

args = parser.parse_args()

video_file = "twitch_nft_demo.mp4"

with open("Xytech.txt", "r") as XF:
    Xytech_file = XF.read().splitlines()
    for currentLine in Xytech_file:
        if currentLine.startswith("Producer: "):
            matched_producer = currentLine.strip().split(':')[1].strip()
        elif currentLine.startswith("Operator: "):
            matched_operator = currentLine.strip().split(':')[1].strip()
        elif currentLine.startswith("Job: "):
            matched_job = currentLine.strip().split(':')[1].strip()
        elif currentLine.startswith("Please "):
            matched_notes = currentLine

matched_frame = []
   
with open("Baselight_export.txt", "r") as BF:
    Baselight_file = BF.read().splitlines()
    #print(file_read_Baselight)
    for currentLine in Baselight_file:
        if not currentLine.strip():
            continue
        #print("Current read: " + currentLine) #/baselightfilesystem1/Dune2/reel1/partA/1920x1080 2 3 4 31 32 33 67 68 69 70 122 123 155 1023 1111 1112 1160 1201 1202 1203 1204 1205 1211 1212 1213 1214 1215
        parseLine = currentLine.split() #['/baselightfilesystem1/Dune2/reel1/partA/1920x1080', '2', '3', '4', '31', '32', '33', '67', '68', '69', '70', '122', '123', '155', '1023', '1111', '1112', '1160', '1201', '1202', '1203', '1204', '1205', '1211', '1212', '1213', '1214', '1215']
        #print(parseLine)
        currentFolder = parseLine.pop(0) #['2', '3', '4', '31', '32', '33', '67', '68', '69', '70', '122', '123', '155', '1023', '1111', '1112', '1160', '1201', '1202', '1203', '1204', '1205', '1211', '1212', '1213', '1214', '1215']
        parseFolder = currentFolder.split("/") #['', 'baselightfilesystem1', 'Dune2', 'reel1', 'partA', '1920x1080']
        parseFolder.pop(1) #['', 'Dune2', 'reel1', 'partA', '1920x1080']
        newFolder = "/".join(parseFolder)  #Dune2/reel1/partA/1920x1080
        #print(newFolder)
        tempStart = 0
        tempLast = 0
        for currentLine in Xytech_file:
            if newFolder in currentLine:
                currentFolder = currentLine.strip()
                #print(currentLine)            

        for number in parseLine:
            if not number.isnumeric() or number == "<err>" or number == "<null>":
                continue
            if tempStart == 0:
                tempStart = int(number)
                continue
            if number == str(int(tempStart) + 1):
                tempLast = int(number)
                continue
            elif number == str(int(tempLast) + 1):
                tempLast = int(number)
                continue
            else:
                if int(tempLast) > 0:
                    matched_frame.append((currentFolder, str(tempStart) + "-" + str(tempLast)))
                else:
                    matched_frame.append((currentFolder, tempStart))
                    
                tempStart = int(number)
                tempLast = 0
        if int(tempLast) > 0:
            matched_frame.append((currentFolder, str(tempStart) + "-" + str(tempLast)))
        else:
            matched_frame.append((currentFolder, tempStart))

#python comp467P3.py --export
if args.EXPORT:
    csv_File = "matched_data.csv"
    with open(csv_File, mode='w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(['Producer ', ' Operator ', ' Job ', ' Notes ' ])
        csv_writer.writerow([matched_producer, matched_operator, matched_job, matched_notes])
        #python comp467P3.py --export --insertX
        if args.INSERTX:
            xytech_data = {
                "Producer": matched_producer,
                "Operator": matched_operator,
                "Job": matched_job,
                "Notes": matched_notes
            }
            mycol2.insert_one(xytech_data)
        csv_writer.writerow([])
        csv_writer.writerow(['Locations: ', ' Frames to fix '])

        for location, frames in matched_frame:
            csv_writer.writerow([location, frames])
            #python comp467P3.py --export --insertBE
            if args.INSERTBE:
                data = mycol1.find({"Frames": args.INSERTBE})
                for frame in str(frames).split(','): 
                    data_df = pd.DataFrame(list(data))
                    data_df = data_df.drop_duplicates(subset=["Frames"]) 
                    baselight_data = {
                        "Location": location,
                        "Frames": frame.strip()  
                    }
                    mycol1.insert_one(baselight_data)
                #python comp467P3.py --export --insertBE --timecode (adds timecode to db)
                if args.TIMECODE:
                    matched_frames_db = []
    
                    for record in mycol1.find():
                        frames = record["Frames"]
                        frame_range = frames.split("-")
                        frame_range_with_timecode = []
                        for frame in frame_range:
                            frame_number = int(frame)
                            total_sec = int(frame_number / fps)
                            remainder = int(frame_number % fps)
                            hours = int(total_sec / 3600)
                            minutes = int((total_sec % 3600) / 60)
                            seconds = int(total_sec % 60)
                            frames = int(remainder)
                            timecode = f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frames:02d}"
                            frame_range_with_timecode.append(timecode)
                        frame_range_string = "-".join(frame_range_with_timecode)
                        matched_frames_db.append(frame_range_string)
        
                    # Updating Baselight_export collection with timecodes
                    for record, timecode in zip(mycol1.find(), matched_frames_db):
                        record["Timecode"] = timecode
                        mycol1.update_one({"_id": record["_id"]}, {"$set": {"Timecode": timecode}})

    print("Uploaded to csv file:", csv_File)

def get_video_duration(video_file):
    result = subprocess.run(['ffprobe', '-i', video_file, '-show_entries', 'format=duration', '-v', 'quiet', '-of', 'csv=p=0'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
    if result.returncode == 0:
        duration_seconds = float(result.stdout)
        milliseconds = int((duration_seconds - int(duration_seconds)) * 1000)
        duration_seconds = int(duration_seconds)
        hours, remainder = divmod(duration_seconds, 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}:{milliseconds:03d}"
    else:
        print("Error: Failed to get video duration.")
        return None
    
def generate_thumbnail(video_file, timecode, output_file, width=96, height=74):
    hours, minutes, seconds, frames = map(int, timecode.split(':'))
    total_seconds = hours * 3600 + minutes * 60 + seconds + frames / 24
    ffmpeg_timecode = f"{int(total_seconds // 3600):02}:{int((total_seconds % 3600) // 60):02}:{total_seconds % 60:.2f}"
    
    subprocess.run(['ffmpeg', '-ss', ffmpeg_timecode, '-i', video_file, '-frames:v', '1', '-vf', f'scale={width}:{height}', output_file], stdout=subprocess.PIPE, stderr=subprocess.PIPE)

    if os.path.isfile(output_file):
        return output_file
    else:
        return None

#python comp467P3.py --process "twitch_nft_demo.mp4"
if args.PROCESS:
    if os.path.isfile(args.PROCESS):
        print("Processing video file:", args.PROCESS)

        video_duration = get_video_duration(args.PROCESS)
        if video_duration is None:
            exit()

        # Converts video duration to milliseconds
        hours, minutes, seconds, milliseconds = map(int, video_duration.split(':'))
        total_duration_ms = (hours * 3600 + minutes * 60 + seconds) * 1000 + milliseconds

        total_frames = total_duration_ms * fps // 1000

        interval = 500
        timecodes = []

        for frame_number in range(1, total_frames + 1):
            total_sec = int(frame_number / fps)
            remainder = int(frame_number % fps)
            hours = int((total_sec / 60) / 60)
            min = int(total_sec / 60)
            sec = int(total_sec % 60)
            frames = int(remainder)

            timecode = f"{hours:02d}:{min:02d}:{sec:02d}:{frames:02d}"
            timecodes.append(timecode)

        print("Video duration:", video_duration)
        print(f"Total frames: {total_frames}")

    else:
        print("Error: Video file not found at:", video_file)
        exit()

#python comp467P3.py --printBE
if args.PRINTBE:
    data = mycol1.find({}, {"Location": 1, "Frames": 1, "Timecode": 1, "_id": 0})

    print("")
    print("From Baselight_export collection 1:")
    print("")

    for row in data:
        for key, value in row.items():
            print(f"{key}: '{value}'")
        print()

#python comp467P3.py --printX
if args.PRINTX:
    data = mycol2.find({}, {"Producer": 1, "Operator": 1, "Job": 1, "Notes": 1, "_id": 0})

    print("")
    print("From Xytech collection 2:")
    print("")

    for row in data:
        for key, value in row.items():
            print(f"{key}: '{value}'")
        print()

#python comp467P3.py --frames
if args.FRAMES:
    matched_frames_db = []
    
    for record in mycol1.find():
        frames = record["Frames"]
        frame_range = frames.split("-")
        frame_range_with_timecode = []
        for frame in frame_range:
            frame_number = int(frame)
            total_sec = int(frame_number / fps)
            remainder = int(frame_number % fps)
            hours = int(total_sec / 3600)
            minutes = int((total_sec % 3600) / 60)
            seconds = int(total_sec % 60)
            frames = int(remainder)
            timecode = f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frames:02d}"
            frame_range_with_timecode.append(timecode)
        frame_range_string = "-".join(frame_range_with_timecode)
        matched_frames_db.append(frame_range_string)

    print("")
    print("Timecode for Baselight_export based on frame range:")    
    for frames in matched_frames_db:
        print(frames)

if args.PROCESSDB:
    if os.path.isfile(args.PROCESSDB):
        print("Processing video file:", args.PROCESSDB)

        video_duration = get_video_duration(args.PROCESSDB)
        if video_duration is None:
            exit()

        # Converts video duration to milliseconds
        hours, minutes, seconds, milliseconds = map(int, video_duration.split(':'))
        total_duration_ms = (hours * 3600 + minutes * 60 + seconds) * 1000 + milliseconds

        total_frames = total_duration_ms * fps // 1000

        interval = 500
        timecodes = []

        for frame_number in range(1, total_frames + 1):
            total_sec = int(frame_number / fps)
            remainder = int(frame_number % fps)
            hours = int((total_sec / 60) / 60)
            min = int(total_sec / 60)
            sec = int(total_sec % 60)
            frames = int(remainder)

            timecode = f"{hours:02d}:{min:02d}:{sec:02d}:{frames:02d}"
            timecodes.append(timecode)

        print("Video duration:", video_duration)
        print(f"Total frames: {total_frames}")

        # Print Baselight export frames and timecodes
        print("\nBaselight export frames and timecodes:")
        for location, frames in matched_frame:
            if isinstance(frames, int):
                frames = f"{frames}-{frames}"
            frame_range = frames.split("-")
            start_frame = int(frame_range[0])
            end_frame = int(frame_range[-1])
            if end_frame > total_frames:
                end_frame = total_frames
            frames = f"{start_frame}-{end_frame}"
            frame_range_with_timecode = []
            total_timecodes = []
            for frame_number in range(start_frame, end_frame + 1):
                if frame_number > total_frames:
                    break
                total_sec = int(frame_number / fps)
                remainder = int(frame_number % fps)
                hours = int(total_sec / 3600)
                minutes = int((total_sec % 3600) / 60)
                seconds = int(total_sec % 60)
                frame_sec = int(remainder)
                timecode = f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frame_sec:02d}"
                total_timecodes.append(timecode)
            if total_timecodes:
                # Only keep the first and last timecodes
                frame_range_with_timecode.extend([total_timecodes[0], total_timecodes[-1]])
                timecode_string = "-".join(frame_range_with_timecode)
                timecode_string = timecode_string.replace("-", "-", 1)
                
                # Checks if frames represent a single frame or a range
                if start_frame != end_frame:
                    timecode_data = {
                        "location": location,
                        "frames": frames,
                        "timecode": timecode_string,
                    }
                    print(f"Location: {location}")
                    print(f"Frames: {frames}")
                    print(f"Timecode: {timecode_string}\n")
                    mycol1.insert_one(timecode_data)

    else:
        print("Error: Video file not found at:", args.PROCESSDB)

#python3 comp467P3.py --exportEX --process "twitch_nft_demo.mp4" 
if args.EXFILE:
    csv_File = "file_contents.csv"
    with open(csv_File, mode='w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)  # Initialize the CSV writer

        def get_video_duration(video_file):
            result = subprocess.run(['ffprobe', '-i', video_file, '-show_entries', 'format=duration', '-v', 'quiet', '-of', 'csv=p=0'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
            if result.returncode == 0:
                duration_seconds = float(result.stdout)
                milliseconds = int((duration_seconds - int(duration_seconds)) * 1000)
                duration_seconds = int(duration_seconds)
                hours, remainder = divmod(duration_seconds, 3600)
                minutes, seconds = divmod(remainder, 60)
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}:{milliseconds:03d}"
            else:
                print("Error: Failed to get video duration.")
                return None

        if args.PROCESS:
            if os.path.isfile(args.PROCESS):
                print("Processing video file:", args.PROCESS)

                video_duration = get_video_duration(args.PROCESS)
                if video_duration is None:
                    exit()

                # Converts video duration to milliseconds
                hours, minutes, seconds, milliseconds = map(int, video_duration.split(':'))
                total_duration_ms = (hours * 3600 + minutes * 60 + seconds) * 1000 + milliseconds

                total_frames = total_duration_ms * fps // 1000

                interval = 500
                timecodes = []

                for frame_number in range(1, total_frames + 1):
                    total_sec = int(frame_number / fps)
                    remainder = int(frame_number % fps)
                    hours = int((total_sec / 60) / 60)
                    min = int(total_sec / 60)
                    sec = int(total_sec % 60)
                    frames = int(remainder)

                    timecode = f"{hours:02d}:{min:02d}:{sec:02d}:{frames:02d}"
                    timecodes.append(timecode)

                    csv_writer.writerow([f"Frame {frame_number}", timecode])

                print("Video duration:", video_duration)
                print(f"Total frames: {total_frames}")
                print("Timecodes for each frame printed to CSV.")

            else:
                print("Error: Video file not found at:", args.PROCESS)

#python comp467P3.py --exportEX --process "twitch_nft_demo.mp4" 
if args.EXFILE:
    csv_File = "file_contents.csv"
    with open(csv_File, mode='w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file) 

        def get_video_duration(video_file):
            result = subprocess.run(['ffprobe', '-i', video_file, '-show_entries', 'format=duration', '-v', 'quiet', '-of', 'csv=p=0'], stdout=subprocess.PIPE, stderr=subprocess.PIPE, universal_newlines=True)
            if result.returncode == 0:
                duration_seconds = float(result.stdout)
                milliseconds = int((duration_seconds - int(duration_seconds)) * 1000)
                duration_seconds = int(duration_seconds)
                hours, remainder = divmod(duration_seconds, 3600)
                minutes, seconds = divmod(remainder, 60)
                return f"{hours:02d}:{minutes:02d}:{seconds:02d}:{milliseconds:03d}"
            else:
                print("Error: Failed to get video duration.")
                return None

        if args.PROCESS:
            if os.path.isfile(args.PROCESS):
                print("Processing video file:", args.PROCESS)

                video_duration = get_video_duration(args.PROCESS)
                if video_duration is None:
                    exit()

                # Converts video duration to milliseconds
                hours, minutes, seconds, milliseconds = map(int, video_duration.split(':'))
                total_duration_ms = (hours * 3600 + minutes * 60 + seconds) * 1000 + milliseconds

                total_frames = total_duration_ms * fps // 1000

                interval = 500
                timecodes = []

                for frame_number in range(1, total_frames + 1):
                    total_sec = int(frame_number / fps)
                    remainder = int(frame_number % fps)
                    hours = int((total_sec / 60) / 60)
                    min = int(total_sec / 60)
                    sec = int(total_sec % 60)
                    frames = int(remainder)

                    timecode = f"{hours:02d}:{min:02d}:{sec:02d}:{frames:02d}"
                    timecodes.append(timecode)

                    csv_writer.writerow([f"Frame {frame_number}", timecode])

                print("Video duration:", video_duration)
                print(f"Total frames: {total_frames}")
                print("Timecodes for each frame printed to CSV.")

            else:
                print("Error: Video file not found at:", args.PROCESS)

#python comp467P3.py --exportCSV --process "twitch_nft_demo.mp4"
if args.CSVFILE:
    csv_File = "videoFile.csv"
    with open(csv_File, mode='w', newline='') as csv_file:
        csv_writer = csv.writer(csv_file)
        csv_writer.writerow(['Location', 'Frames to fix', 'Timecode', 'Thumbnail'])
        if args.PROCESS:
            if os.path.isfile(args.PROCESS):

                video_duration = get_video_duration(args.PROCESS)
                if video_duration is None:
                    exit()

                # Converts video duration to milliseconds
                hours, minutes, seconds, milliseconds = map(int, video_duration.split(':'))
                total_duration_ms = (hours * 3600 + minutes * 60 + seconds) * 1000 + milliseconds

                total_frames = total_duration_ms * fps // 1000

                interval = 500
                timecodes = []

                for frame_number in range(1, total_frames + 1):
                    total_sec = int(frame_number / fps)
                    remainder = int(frame_number % fps)
                    hours = int((total_sec / 60) / 60)
                    min = int(total_sec / 60)
                    sec = int(total_sec % 60)
                    frames = int(remainder)

                    timecode = f"{hours:02d}:{min:02d}:{sec:02d}:{frames:02d}"
                    timecodes.append(timecode)

                # Print Baselight export frames and timecodes
                for location, frames in matched_frame:
                    if isinstance(frames, int):
                        frames = f"{frames}-{frames}"
                    frame_range = frames.split("-")
                    start_frame = int(frame_range[0])
                    end_frame = int(frame_range[-1])
                    if end_frame > total_frames:
                        end_frame = total_frames
                    frames = f"{start_frame}-{end_frame}"
                    frame_range_with_timecode = []
                    total_timecodes = []
                    for frame_number in range(start_frame, end_frame + 1):
                        if frame_number > total_frames:
                            break
                        total_sec = int(frame_number / fps)
                        remainder = int(frame_number % fps)
                        hours = int(total_sec / 3600)
                        minutes = int((total_sec % 3600) / 60)
                        seconds = int(total_sec % 60)
                        frame_sec = int(remainder)
                        timecode = f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frame_sec:02d}"
                        total_timecodes.append(timecode)
                    if total_timecodes:
                        # Only keep the first and last timecodes
                        frame_range_with_timecode.extend([total_timecodes[0], total_timecodes[-1]])
                        timecode_string = "-".join(frame_range_with_timecode)
                        timecode_string = timecode_string.replace("-", "-", 1)

                        # Calculates the middle-most frame
                        mid_frame = (start_frame + end_frame) // 2
                        total_sec = int(mid_frame / fps)
                        remainder = int(mid_frame % fps)
                        hours = int(total_sec / 3600)
                        minutes = int((total_sec % 3600) / 60)
                        seconds = int(total_sec % 60)
                        frame_sec = int(remainder)
                        middle_timecode = f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frame_sec:02d}"

                        # Generates thumbnail
                        thumbnail_path = f"thumbnail_{start_frame}_{end_frame}.jpg"
                        thumbnail = generate_thumbnail(args.PROCESS, middle_timecode, thumbnail_path)
                
                        # Check if frames represent a single frame or a range
                        if start_frame != end_frame:
                            timecode_data = {
                                "location": location,
                                "frames": frames,
                                "timecode": timecode_string,
                                "thumbnail": thumbnail
                            }
                            csv_writer.writerow([location, frames, timecode_string, thumbnail])

            else:
                print("Error: Video file not found at:", args.PROCESS)

def timecode_to_seconds(timecode, fps):
    # Converting timecode to seconds
    hours, minutes, seconds, frames = map(int, timecode.split(':'))
    total_seconds = hours * 3600 + minutes * 60 + seconds + frames / float(fps)
    return total_seconds

def seconds_to_timecode(seconds):
    # Convertitng seconds to timecode in HH:MM:SS.ms format
    return str(datetime.timedelta(seconds=seconds))

def extract_clip(video_file, start_timecode, end_timecode, output_file): 
    # Converts start and end timecodes to seconds
    start_seconds = timecode_to_seconds(start_timecode, fps)
    end_seconds = timecode_to_seconds(end_timecode, fps)

    # Converts seconds to HH:MM:SS.ms format
    ffmpeg_start_time = seconds_to_timecode(start_seconds)
    ffmpeg_end_time = seconds_to_timecode(end_seconds)

    try:
        subprocess.run(['ffmpeg', '-y', '-ss', ffmpeg_start_time, '-i', video_file, '-to', ffmpeg_end_time,
                '-c:v', 'libx264', '-crf', '23', '-preset', 'medium', '-c:a', 'aac', '-b:a', '128k',
                '-movflags', 'faststart',
                output_file], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=True)

        if os.path.isfile(output_file):
            return output_file
        else:
            print("Error: Output file does not exist.")
            return None
    except subprocess.CalledProcessError as e:
        print("Error occurred while extracting clip:", e)
        return None

# python comp467P3.py --output
if args.XLS:
    xlsx_file = "matched_data.xlsx"

    workbook = excel.Workbook(xlsx_file)
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, 'Producer')
    worksheet.write(0, 1, 'Operator')
    worksheet.write(0, 2, 'Job')
    worksheet.write(0, 3, 'Notes')

    worksheet.write(1, 0, matched_producer)
    worksheet.write(1, 1, matched_operator)
    worksheet.write(1, 2, matched_job)
    worksheet.write(1, 3, matched_notes)

    worksheet.write(2, 0, '') 

    worksheet.write(3, 0, 'Locations')
    worksheet.write(3, 1, 'Frames to fix')
    worksheet.write(3, 2, "Timecode")
    worksheet.write(3, 3, "Thumbnail")

    if args.PROCESS:
        if os.path.isfile(args.PROCESS):

            video_duration = get_video_duration(args.PROCESS)
            if video_duration is None:
                exit()

            hours, minutes, seconds, milliseconds = map(int, video_duration.split(':'))
            total_duration_ms = (hours * 3600 + minutes * 60 + seconds) * 1000 + milliseconds

            total_frames = total_duration_ms * fps // 1000

            timecodes = []

            for frame_number in range(1, total_frames + 1):
                total_sec = int(frame_number / fps)
                remainder = int(frame_number % fps)
                hours = int(total_sec / 3600)
                minutes = int((total_sec % 3600) / 60)
                seconds = int(total_sec % 60)
                frame_sec = int(remainder)

                timecode = f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frame_sec:02d}"
                timecodes.append(timecode)

            for idx, (location, frames) in enumerate(matched_frame, start=4):
                if isinstance(frames, int):
                    frames = f"{frames}-{frames}"
                frame_range = frames.split("-")
                start_frame = int(frame_range[0])
                end_frame = int(frame_range[-1])
                if end_frame > total_frames:
                    end_frame = total_frames
                frames = f"{start_frame}-{end_frame}"
                frame_range_with_timecode = []
                total_timecodes = []

                for frame_number in range(start_frame, end_frame + 1):
                    if start_frame == end_frame:
                        continue
                    if frame_number > total_frames:
                        break
                    total_sec = int(frame_number / fps)
                    remainder = int(frame_number % fps)
                    hours = int(total_sec / 3600)
                    minutes = int((total_sec % 3600) / 60)
                    seconds = int(total_sec % 60)
                    frame_sec = int(remainder)
                    timecode = f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frame_sec:02d}"
                    total_timecodes.append(timecode)

                if total_timecodes:
                    frame_range_with_timecode.extend([total_timecodes[0], total_timecodes[-1]])
                    timecode_string = "-".join(frame_range_with_timecode)
                    timecode_string = timecode_string.replace("-", "-", 1)

                    # Calculates the middle-most frame
                    mid_frame = (start_frame + end_frame) // 2
                    total_sec = int(mid_frame / fps)
                    remainder = int(mid_frame % fps)
                    hours = int(total_sec / 3600)
                    minutes = int((total_sec % 3600) / 60)
                    seconds = int(total_sec % 60)
                    frame_sec = int(remainder)
                    middle_timecode = f"{hours:02d}:{minutes:02d}:{seconds:02d}:{frame_sec:02d}"

                    start_timecode = total_timecodes[0]
                    end_timecode = total_timecodes[-1]
                    
                    # Generates thumbnail
                    thumbnail_path = f"thumbnail_{start_frame}_{end_frame}.jpg"
                    thumbnail = generate_thumbnail(args.PROCESS, middle_timecode, thumbnail_path)
                    # Generates clip
                    clip_filename = f"clip_{start_frame}_{end_frame}.mp4"
                    clip_path = extract_clip(args.PROCESS, start_timecode, end_timecode, clip_filename)

                    if clip_path:
                        print(f"Location: {location}")
                        print(f"Frames: {frames}")
                        print(f"Timecode: {timecode_string}")
                        print(f"Clip: {clip_path}\n")

                        if start_frame != end_frame:
                            timecode_data = {
                                "location": location,
                                "frames": frames,
                                "timecode": timecode_string,
                                "thumbnail": thumbnail
                            }
                            worksheet.write(idx, 0, location)
                            worksheet.write(idx, 1, frames)
                            worksheet.write(idx, 2, timecode_string)
                            worksheet.insert_image(idx, 3, thumbnail)

                        if args.FRAMEIO:
                            frameio_token = 'TOKEN'
                            asset_id = 'ID'

                            client = FrameioClient(token=frameio_token)

                            if os.path.isfile(clip_path):
                                asset = client.assets.upload(destination_id=asset_id, filepath=clip_path)
                            else:
                                print("Error: Video clip file not found at:", clip_path)

        else:
            print("Error: Video file not found at:", args.PROCESS)

    workbook.close()

    print("Uploaded to xlsx file:", xlsx_file)