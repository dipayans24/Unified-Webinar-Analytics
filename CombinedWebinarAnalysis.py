import io
import os
import streamlit as st
import pandas as pd, re
from pandas.errors import ParserWarning
import warnings
warnings.filterwarnings("ignore", category=ParserWarning)
from datetime import datetime
import csv
import tempfile, openpyxl


def getNumber(number):
    try:
        mobile =  re.findall(r'\d+', str(number))[0]
        if len(mobile) == 11 and mobile[0] == "0":
            return "91"+str(mobile[1:])
        elif len(mobile)==10:
            return "91"+str(mobile)
            #return str(mobile)
        else:
            return str(mobile)
    except:
        return pd.NA

def getCleanPhone(number):
    try:
      return  re.findall(r'\d+', str(number))[0]
    except:
      return pd.NA

def getAttendanceFormat(filePath, AttendanceTimeThreshold = 0):
  #fileName = filePath.split(".")[0].replace("attendee_", "")
  #folderName = r"C:\Users\Dipayan\Desktop\New folder (5)"
  #While loop for reading the file untill the parser error is not encountered.
  cnt = 0
  while cnt != None:
      try:
          attendance = pd.read_csv(filePath, sep="," , index_col=False ,  skiprows=cnt )
          cnt = None
      except:
          cnt = cnt+1


  #print(filePath)
  #drops all the join time rows which are blank.
  attendance = attendance[(attendance["Join Time"] != "--")].dropna(subset=["Join Time"])

  #drops all the irrelevant rows
  attendance = attendance.where(lambda x : x["Attended"] == "Yes", other=pd.NA).dropna(subset=["Attended"])

  #Change the datatype of Join Time and Leave Time
  attendance[["Join Time", "Leave Time"]] = attendance.loc[:, ["Join Time", "Leave Time"]].apply(lambda x : pd.to_datetime(x, format="%m/%d/%Y %I:%M:%S %p"))

  attendanceDate = attendance["Join Time"].dt.date.unique()[0]

  #fileNamePart = attendance["Join Time"].dt.strftime("%d%b%Y").unique()[0]

  #Extracts only the digits from the Phone column
  attendance["OriginalNumber"] = attendance.Phone.apply(lambda x: getCleanPhone(x)).astype(float, errors="ignore")

  attendance["Phone"] = attendance["Phone"].apply(lambda x: getNumber(x)).astype(float, errors="ignore")

  #Change the datatype to float for the Time in session column
  attendance["Time in Session (minutes)"] = attendance["Time in Session (minutes)"].fillna(0).astype(float, errors="ignore")

  df = pd.read_csv(filePath, sep=",",  skiprows=2, on_bad_lines="skip")
  TopicName = re.sub(r"[^a-zA-Z0-9\s]", "", df["Topic"].unique()[0]) 
  WebinarID =  df.loc[:, "Webinar ID"].unique()[0].replace(" ", "") 
  attendance["WebinarID"] = int(WebinarID)

  #drops all the irrelevant rows
  #attendance = attendance.where(lambda x : x["Attended"] == "Yes", other=pd.NA).dropna(subset=["Attended"])
  
  #Group by using Email to get the total duration for that user
  Duration  = attendance.groupby(by="Email", as_index=False).agg( SessionDuration = ("Time in Session (minutes)", "sum"), UserName =  ("User Name (Original Name)", "max"))

  #Drop the duplicate rows using the Email column
  cleanAttendance = attendance.drop_duplicates(subset="Email", keep="first").loc[:, ["Email", "Phone", "OriginalNumber", "WebinarID"]]

  #Merges the both the dataframe
  final = cleanAttendance.merge(Duration, how="left", left_on="Email", right_on="Email")

  #Change the datatype to float for the Phone column
  final.Phone = final.Phone.fillna(0).astype(float, errors="ignore")

  #Filters the data by the value mentioned in the AttendanceTimeThreshold threshold
  final = final[final["SessionDuration"] > AttendanceTimeThreshold]

  #Sorts the dataframe using the Time in session(minutes)
  final.sort_values(by="SessionDuration", ascending=False, inplace=True)

  final["Date"] = None
  final["Date"] = final["Date"].apply(lambda x: attendanceDate if x is None else x)

  final = final.loc[: , ["Date",  "WebinarID", "UserName", "Email", "Phone", "OriginalNumber", "SessionDuration"]]

  return final, TopicName, WebinarID


def formatChat(chats):
    _INVALID_CHARS = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')
    timestamps, comments = [], []
    current_comment = None
    current_lines = []

    for raw in chats:
        line = _INVALID_CHARS.sub('', raw.decode("utf-8"))

        if any(kw in line for kw in ("panelists:", " Everyone:", "(direct message)")):
            # Save previous comment with all its accumulated lines
            if current_comment is not None:
                timestamps.append(current_comment)
                comments.append("\n".join(current_lines).strip() if current_lines else None)
            
            current_comment = line
            current_lines = []  # Reset lines for new message
        else:
            if current_comment is not None:
                # Accumulate continuation lines instead of closing immediately
                if line.strip():
                    current_lines.append(line.strip())

    # Don't forget the last message
    if current_comment is not None:
        timestamps.append(current_comment)
        comments.append("\n".join(current_lines).strip() if current_lines else None)

    data = pd.DataFrame({"TimeStamp": timestamps, "Comments": comments})

    ts_split = data["TimeStamp"].str.split(" ", n=1, expand=True)
    info_split = ts_split[1].str.split(" to ", n=1, expand=True)

    data["Time"] = ts_split[0]
    data["From"] = info_split[0].str.replace("From", "", regex=False).str.strip()
    data["To"]   = (
        info_split[1]
        .str.replace(":", "", regex=False)
        .str.strip()
        .str.replace(", [Hh]osts and panelists", "", regex=True)
    )
    data["Comments"] = data["Comments"].str.strip()
    data = data[["Time", "From", "To", "Comments"]]

    ChatAnalysis = data.groupby(by="From", as_index=False).agg(
                    UniqueCount=("Comments", "nunique"),
                    TotalCount=("Comments", "count")
                )
    ChatAnalysis["SpamPercentage"] = ChatAnalysis.apply(
    lambda x: round(100 - ((x["UniqueCount"] * 100) / x["TotalCount"]), 2), axis=1
    )
    ChatAnalysis = ChatAnalysis[ChatAnalysis["UniqueCount"] > 1]
    ChatAnalysis["Replies"] = ChatAnalysis["From"].apply(
    lambda x: data[data["To"] == x]["To"].count()
    )
    ChatAnalysis.sort_values(by="SpamPercentage", ascending=False, inplace=True)

    return ChatAnalysis, data

def getPollDetails(pollPath):

    #Get the pollDetails
    pollDetails = pd.read_csv(pollPath, sep=",", skiprows=1, on_bad_lines="skip", nrows=1)
    WebinarID = pollDetails["Meeting/Webinar ID"].values[0]
    #TopicName = pollDetails["Meeting Topic"].values[0]
    #SessionDate = pollDetails["Actual Start Time"].apply(lambda x : pd.to_datetime(x)).dt.date.values[0]

    #Get the list of polls
    pollList = pd.read_csv(pollPath, sep=",", skiprows=5, on_bad_lines="skip")
    pollList = pollList.dropna(subset=pollList.columns[1:], how="all").iloc[:, 1:]
    pollList["rnk"] = pollList["Responses"].rank(method="dense", ascending=False)
    pollList = pollList[pollList["rnk"]<3]
    df = []

    startPosition = []
    stopPosition = [i for i in pollList["Responses"].tolist()]
    variableNames = set()

    with open(pollPath, "r", encoding="utf-8") as f:
        reader = csv.reader(f)
        for logical_row, row in enumerate(reader):
            for pollName in pollList["Poll Name"].tolist():
                variableNames.add(re.sub(r"\W","", pollName)) #re.sub(r"\W","", pollName)
                if row and row[0].strip() == pollName.strip():
                    startPosition.append(logical_row)

    start_stop = {i:j for i,j in zip(startPosition, stopPosition)}

    for names, start, end in zip(variableNames, start_stop.keys(), start_stop.values()):
        globals()[names] = pd.read_csv(pollPath, sep=",", skiprows=start+1, index_col=False, encoding="utf-8" , nrows=end)
        globals()[names] = globals()[names].iloc[:, 2:]
        globals()[names].rename(columns={"Email Address":"Email"}, inplace =True)
        df.append(globals()[names])

    return df, WebinarID

def save_upload(uploaded_file):
    """Save a Streamlit UploadedFile to a temp file and return the path."""
    suffix = os.path.splitext(uploaded_file.name)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.read())
    tmp.flush()
    tmp.close()
    return tmp.name

def getRegistration(filePath):
    registrationSummary = pd.read_csv(filePath, sep=",", skiprows=2, nrows=1)
    registration = pd.read_csv(filePath, sep=",",skiprows=5, skip_blank_lines=True,)
    registration["OriginalNumber"] = registration.Phone.apply(lambda x: getCleanPhone(x)).astype(float, errors="ignore")
    registration["Phone"] = registration["Phone"].apply(lambda x: getNumber(x)).astype(float, errors="ignore")
    registration["UserName"] = registration["First Name"].fillna('') + " " + registration["Last Name"].fillna('')
    registrationID = registrationSummary["ID"].unique()[0].replace(" ", "")
    registration["WebinarID"] = int(registrationID)
    registration["Date" ] = pd.to_datetime(registrationSummary["Scheduled Time"]).dt.date.values[0]
    registration = registration[registration["Approval Status"].str.lower() == "approved"]
    
    registration = registration.loc[:, ["Date",  "WebinarID", 'UserName', 'First Name', 'Last Name', 'Email', 'Registration Time',
       'Approval Status', 'Phone', 'OriginalNumber']]
    return  registration, registrationID
    
@st.dialog("Error Alert")
def raiseError(text):   
    st.error(text)

def generateOutput(dataFrames, sheet_names):
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer) as f:
        for df, sheet in zip(dataFrames, sheet_names):
            if df is not None:
                df.to_excel(f, index=False, sheet_name=sheet)

    output_buffer.seek(0)
    return output_buffer
    
def processFiles(filePath, chat_file_path = None, pollPath = None, include_raw_chat = False, isRegistration = False):
    chats = []

    if not isRegistration:
        attendance, TopicName, AttendanceWebinarID = getAttendanceFormat(filePath=filePath)
    else:
        attendance, TopicName, AttendanceWebinarID = getRegistration(filePath=filePath)

    RawChat = None
    if chat_file_path is not None:

        with tempfile.TemporaryDirectory() as tmpdir:
            for uploaded_file in chat_file_path:
                file_path = os.path.join(tmpdir, uploaded_file.name)
                with open(file_path, "wb") as f:
                    f.write(uploaded_file.read())
                with open(file_path, "rb+") as f:
                    chats.extend(f.readlines())
        ChatAnalysis, RawChat= formatChat(chats)
        mergedAttendance = attendance.merge(ChatAnalysis, left_on = "UserName", right_on="From", how="left")
    
    else:
        mergedAttendance = attendance

    if pollPath is not None:
        pollPath = save_upload(pollPath)
        Polls,  PollWebinarID  = getPollDetails(pollPath)
        if int(AttendanceWebinarID) == int(PollWebinarID):
            for poll in Polls:
                mergedAttendance = mergedAttendance.merge(poll, left_on="Email", right_on="Email", how="left")

            return generateOutput([mergedAttendance]+[RawChat if RawChat is not None and include_raw_chat is True else None], [TopicName[:30], "Chat"]), AttendanceWebinarID

        elif int(AttendanceWebinarID) != int(PollWebinarID):
            raiseError("There is a mismatch between Attendee File and Poll File.")
            return None, AttendanceWebinarID
        
    else:

        return generateOutput([mergedAttendance]+[RawChat if RawChat is not None and include_raw_chat is True else None], [TopicName[:30], "Chat"]), AttendanceWebinarID


st.set_page_config(
    page_title="Merge Zoom Files",
    page_icon="🎦",
    layout="wide"
)
 
st.header("🎦Merge Zoom Files")

st.info('''Required: Zoom Attendee File\n
Optional (recommended for a complete report):\n
● Zoom Chat File\n
● Zoom Poll File\n
Upload the attendee file and at least one of the optional files to generate a comprehensive report.''')

filePath = st.file_uploader("Upload the Attendee File", accept_multiple_files=False, type=["csv"], width="stretch")
isRegistrationFile = st.checkbox("This is a Zoom Registration File.")


chat_file_path = st.file_uploader("Optional- Upload the Chat File(s). Multple files supported.", accept_multiple_files=True, type=["txt"], width="stretch")

if chat_file_path:
    include_raw_chat = st.checkbox("Includes raw chat file if selected.")
else:
    include_raw_chat = False

pollPath = st.file_uploader("Optional- Upload the Poll Report",accept_multiple_files=False, type=["csv"], width="stretch")


button  = st.button("Process Files", type="primary", key="button")
 
if button:
    if filePath is not None and (chat_file_path is not None or pollPath is not None):   
        try:
            with st.spinner("Getting Ready..."):
                if not chat_file_path:
                    chat_file_path = None
                if not pollPath:
                    pollPath = None
    
                df, WebinarID = processFiles(save_upload(filePath), chat_file_path, pollPath, include_raw_chat, isRegistrationFile)
                
                output_filename = rf"attendee_{str(WebinarID)}_Formatted.xlsx"
    
                if df is not None:  
                    st.download_button(
                        label=" Download Excel Report",
                        data=df,
                        file_name= output_filename ,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary", 
                        icon="⬇️" 
                    )
        except Exception as e:
            raiseError("Please upload zoom attendee/registration/chat/poll files only.")
            
    elif filePath is None:
        raiseError("Please upload the zoom attendee file.")
