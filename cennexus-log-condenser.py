# standard imports
from openpyxl import Workbook, load_workbook, utils
from tqdm import tqdm
import sys, getopt, csv, os

# defined constants
DEBUG = False # set to False for regular usage
DEBUG_ROWS = 100
DESCRIPTION_COL_NUM = 3
SEND_STRING = 'Send: <STX>2M|'
RECEIVE_STRING = 'Receive: <STX>3O|'
LOGIN_STRING = '2M|1|101'
STORAGE_STRING = '2M|1|103|'
DEFAULT_SAVE = 'parsed.xlsx'
USAGE = """usage: cennexus-log-condenser.py -i <inputfile> -o <outputfile> 
alt-usage: cennexus-log-condenser.py -d <directory>"""

def convert_csv(filename):
  """Converts a csv to xlsx

  Basically just copies all the rows into a new workbook

  @return a string for the new file (foo.xlsx)
  """
  wb = Workbook()
  ws = wb.active
  with open(filename, 'r') as f:
    for row in csv.reader(f):    
      ws.append(row)
  f.close()
  index = filename.index('.csv')
  filename = filename[:index] + '.xlsx'
  wb.save(filename)
  wb.close()
  return filename


def parse_xlsx(inputfile, outputfile):
  """Parses an xlsx to get only the host Order and Manufacturer messages

  Includes a progress bar.

  Some extra processing to get the date and time separated.

  If the original file was from a .csv, it will delete the temporary .xlsx file.

  @param inputfile the file to process
  @param outputfile the desired name of the output file
  """
  # open the workbooks and get the active sheets
  print('Loading workbook: {}'.format(inputfile))
  wb = load_workbook(filename=inputfile, read_only=True)
  ws = wb.active
  nb = Workbook(write_only=True)
  ns = nb.create_sheet()
  row_count = ws.max_row
  if DEBUG:
    if (wb.epoch == utils.datetime.CALENDAR_WINDOWS_1900):
      print('Workbook using the 1900 date system')
    else:
      print('Workbook not using the 1900 date system')
  print('Loaded Workbook!')

  # loop through each row and write it to the new file if it is a valid message
  print('Processing data... Please be patient')
  ns.append(['Timestamp', 'Message', 'isReceive', 'isSend',
    'isLogin', 'isStorage', 'Date', 'Hour', 'Minute', 'Time'])

  # for debugging, limit how many rows are processed
  if DEBUG:
    i = 0
    row_count = DEBUG_ROWS

  # using a progress bar here
  with tqdm(total=row_count, ascii=True, desc='Working...') as pbar:
    for row in ws.rows:
      if DEBUG:
        i = i + 1
        if i == DEBUG_ROWS + 1:
          break
      pbar.update(1)
      message = row[DESCRIPTION_COL_NUM].value
      if message is None: # just in case a row is empty
        continue
      if DEBUG:
        print(message)
        print('Starts with SEND?: {}'.format(message.startswith(SEND_STRING)))
        print('Starts with RECEIVE?: {}'.format(message.startswith(RECEIVE_STRING)))

      # check for valid messages and process message counts and date/time
      isSend = 0
      isReceive = 0
      isLogin = 0
      isStorage = 0
      if message.startswith(RECEIVE_STRING):
        isReceive = 1
      elif message.startswith(SEND_STRING):
        isSend = 1
        if LOGIN_STRING in message:
          isLogin = 1
        elif STORAGE_STRING in message:
          isStorage = 1
         
      # keep valid messages or move on
      if (isSend == 1 or isReceive == 1):

        # get the various date components
        timestamp = row[0].value
        split_timestamp = timestamp.split(' ')
        date = split_timestamp[0]
        time_string = split_timestamp[1]
        time_array = time_string.split(':')
        hour = int(time_array[0])
        if split_timestamp[2] == 'PM':
          hour = hour + 12
        minute = int(time_array[1])
        if int(minute) < 10:
          time = str(hour) + ':0' + str(minute) 
        else:
          time = str(hour) + ':' + str(minute)

        if DEBUG:
          print('Valid message, appending to new worksheet...')
        ns.append([timestamp, message, isReceive, isSend, 
          isLogin, isStorage, date, hour, minute, time])
        continue
        if DEBUG:
          print('SEND/RECEIVE not found, skipping row')

  # close workbook and progress bar since they are done
  pbar.close()
  wb.close()
  print('Data parse complete!')

  # save the new data in a fresh workbook
  print('Saving to ' + outputfile + '...')
  nb.save(outputfile)
  nb.close()
  if isCSV:
    os.remove(inputfile)
  print('Data saved!')
  return

# main method
def main(argv):
   """
   This program will read in an .xlsx or .csv and parse out the Order or
   Manufacturer messages.  It also adds some headers for easier pivot tableing.
   Do note the dependencies.
   """
   print('Cennexus Log Host Message Condenser')
   if DEBUG:
      print('Argument count: ' + str(len(argv)))
      print('Arguments: ' + str(argv))

   input_file = ''
   output_file = DEFAULT_SAVE
   dir_source = ''
   global isCSV
   isCSV = False

   # get the arguments for input and output file
   try:
       opts, args = getopt.getopt(argv, 'hi:o:d:',['input=', 'output=', 'dir='])

   except getopt.GetoptError:
      print(USAGE)
      sys.exit(2)

   for opt, arg in opts:
      if DEBUG:
         print('Parsing option {} and argument {}'.format(opt, arg))
      if opt == '-h':
         print(USAGE)
         sys.exit(2)
      elif opt in ('-d', '--dir'):
         dir_source = arg
      elif opt in ('-i', '--input'):
         input_file = arg
      elif opt in ('-o', '--output'):
         output_file = arg

   if DEBUG:
      print('Input file: ' + input_file)
      print('Output file: ' + output_file)
      print('Directory source: ' + dir_source)

   # quit if there is no input file; output file has a default filename
   if input_file == '' and dir_source == '':
      print(USAGE)
      sys.exit(2)

   # if the input is .csv, convert to .xlsx
   if input_file.endswith('.csv'):
       print('This is a .csv file, converting to .xlsx...')
       input_file = convert_csv(input_file)
       isCSV = True


   # convert the file
   if dir_source == '':
       parse_xlsx(input_file, output_file)   
       print('Bye bye')
       sys.exit(0)

   else:
       print('Directory processing not implemented')
       sys.exit(1)

# runs main
if __name__ == '__main__':
   main(sys.argv[1:])

