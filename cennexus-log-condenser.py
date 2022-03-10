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
DEFAULT_SAVE = 'parsed.xlsx'
USAGE = 'cennexus-log-condenser.py -i <inputfile> -o <outputfile>'

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

   # get the arguments for input and output file
   try:
      opts, args = getopt.getopt(argv, 'hi:o:',['ifile=', 'ofile='])

   except getopt.GetoptError:
      print(USAGE)
      sys.exit(2)

   for opt, arg in opts:
      if DEBUG:
         print('Parsing option {} and argument {}'.format(opt, arg))
      if opt == '-h':
         print(USAGE)
         sys.exit(2)
      elif opt in ('-i', '--ifile'):
         input_file = arg
      elif opt in ('-o', '--ofile'):
         output_file = arg

   if DEBUG:
      print('Input file: ' + input_file)
      print('Output file: ' + output_file)

   # quit if there is no input file; output file has a default filename
   if input_file == '':
      print(USAGE)
      sys.exit(2)

   # if the input is .csv, convert to .xlsx
   if input_file.endswith('.csv'):
       print('This is a .csv file, converting to .xlsx...')
       csv_wb = Workbook()
       csv_ws = csv_wb.active
       with open(input_file, 'r') as f:
           for row in csv.reader(f):
               csv_ws.append(row)
       f.close()
       index = input_file.index('.csv')
       input_file = input_file[:index] + '.xlsx'
       csv_wb.save(input_file)
       csv_wb.close()


   # open the workbooks and get the active sheets
   print('Loading workbook: {}'.format(input_file))
   wb = load_workbook(filename=input_file, read_only=True)
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
   ns.append(['Timestamp', 'Message', 'isSend', 'isReceive', 'Date', 'Hour', 'Minute', 'Time'])

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
         if message is None:
             continue
         if DEBUG:
            print(message)
            print('Starts with SEND?: {}'.format(message.startswith(SEND_STRING)))
            print('Starts with RECEIVE?: {}'.format(message.startswith(RECEIVE_STRING)))

         # check for valid messages
         isSend = 0
         isReceive = 0
         if message.startswith(SEND_STRING):
            isSend = 1
         elif message.startswith(RECEIVE_STRING):
            isReceive = 1
         
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
            ns.append([timestamp, message, isSend, isReceive, date, hour, minute, time])
            continue
         if DEBUG:
            print('SEND/RECEIVE not found, skipping row')

   # close workbook and progress bar since they are done
   pbar.close()
   wb.close()
   print('Data parse complete!')

   # save the new data in a fresh workbook
   print('Saving to ' + output_file + '...')
   nb.save(output_file)
   nb.close()
   print('Data saved!')
   print('Bye bye')

# runs main
if __name__ == '__main__':
   main(sys.argv[1:])
