#!/usr/bin/python

# Author: Saicharan.G
# Code: WhatsApp automation

# https://web.whatsapp.com/send?phone=919490123292&text=hi

try:
      from twilio.rest import Client
      from tkinter import *
      from tkinter import filedialog
      from tkinter import messagebox
      from threading import Thread
      import xlrd
      import xlwt
      import os
except ImportError:
      print( '\n>>> You dont\'ve required modules\n' )
      quit()

def show_message( message, die = False ):
      messagebox.showinfo( 'Information', message )
      if die:
            exit( 1 )

file = label = thread = window = None

def update_status( status ):
      label.config( text = 'Status :     ' + status )
      window.update()

class WhatsApp:
      connected = None
      whatsapp = None
      TWILIO_ACCOUNT_ID = None
      TWILIO_AUTHENTICATION_CODE = None

      def __init__( self ):
            update_status( 'Connecting to TWILIO server...' )
            try:
                  with open( 'config.key', encoding = 'utf-8-sig' ) as fd:
                        temp = fd.readline().strip().split()
                        self.TWILIO_ACCOUNT_ID = temp[0]
                        self.TWILIO_AUTHENTICATION_CODE = temp[1][:-1] if temp[1][-1] == '\n' else temp[1]
            except FileNotFoundError:
                  show_message( '\"config.key\" file is not found' )
                  return
            except:
                  show_message( 'We got error while reading \"config.key\" file.\nCheck it' )
                  return
            print( self.TWILIO_ACCOUNT_ID + ' ' + self.TWILIO_AUTHENTICATION_CODE )
            try:
                  self.whatsapp = Client( self.TWILIO_ACCOUNT_ID, self.TWILIO_AUTHENTICATION_CODE )
                  self.connected = True
            except:
                  show_message( 'We got an error while connecting to whatsapp server\nTry agin' )

      def send_whatsapp( self, reciever_number, message ):
            try:
                  self.whatsapp.messages.create( body = message, from_ = 'whatsapp:+14155238886', to = 'whatsapp:+' + reciever_number )
            except Exception as e:
                  return True

      def __del__( self ):
            update_status( 'No process is running' )

def start_process():
      object = WhatsApp()
      if not object.connected:
            update_status( 'No process is running' )
            return

      try:
            read_book = xlrd.open_workbook( file )
      except:
            show_message( 'We got an error when reading ' + file + '\nCheck that file' )
            return

      write_book = xlwt.Workbook()
      write_sheet = write_book.add_sheet( 'Sheet 1' )
      write_sheet.write( 0, 0, 'Name' )
      write_sheet.write( 0, 1, 'Email' )
      write_sheet.write( 0, 2, 'Candidate ID' )
      write_sheet.write( 0, 3, 'Mobile no' )
      count = 0
      for index in range( len(read_book.sheet_names()) ):
            read_sheet = read_book.sheet_by_index( index )
            count += read_sheet.nrows - 1
      count = str( count )
      update_status( 'Sended msgs: 0/' + count )
      for index in range( len(read_book.sheet_names()) ):
            read_sheet = read_book.sheet_by_index( index )
            i = 1
            not_sended = sended = 0
            while i < read_sheet.nrows:
                  try:
                        list = read_sheet.row_values( i )
                        reciever_name, reciever_mail, c_id, reciever_number = list
                        reciever_number = str(int( reciever_number ))
                        message = 'Hello ' + reciever_name + ',\n\nWe wish to inform you that your payment is due.\nPlease pay using this link.\nYour candidate ID : ' + str(int(c_id)) + '\n\nThanks and Regards.'
                  except Exception as exception:
                        show_message( str(i) + ' line data is incorrect check it.\n' + exception.message, True )
                  if not object.send_whatsapp( reciever_number, message ):
                        sended += 1
                        update_status( 'Sended msgs: ' + str(sended) + '/' + count )
                  else:
                        not_sended += 1
                        for j in range( 4 ):
                              write_sheet.write( not_sended, j, list[j] )
                  i += 1
      final_output = 'Number of users: ' + count + '\nWhatsApp msgs sended: ' + str(sended) + '\nWhatsApp msgs not sended: ' + str(not_sended)
      if not_sended:
            head, tail = os.path.split( file )
            output = head + '/erros_' + tail
            write_book.save( output )
            final_output += '\n\nFailed msgs list created in \"' + output + '\"'
      show_message( final_output )

def main():
      global label, window
      window = Tk()
      window.title( 'WhatsApp Automation' )
      window.geometry( '405x120' )
      window.resizable( False, False )
      window.eval( 'tk::PlaceWindow . center' )

      def browse_files( event = None ):
            global file
            file = filedialog.askopenfilename( initialdir = ".", title = "Select .xlsx file", filetypes = (('xlsx files', '*.xlsx'),) )
            entry.insert( 0, file )

      def create_thread():
            global thread
            if not file:
                  show_message( 'Xlsx file is not selected' )
                  return
            if thread and thread.isAlive():
                  show_message( 'Process was already started' )
                  return
            thread = Thread( target = start_process )
            thread.start()

      entry = Entry( window )
      entry.insert( 0, 'Select Xlsx file path' )
      entry.grid( row = 1, column = 0, padx = (25, 0), pady = 15, ipady = 4, ipadx = 25 )
      entry.bind( '<Button-1>', browse_files )

      button1 = Button( window, text = "Select Xlsx file", command = browse_files )
      button1.grid( row = 1, column = 1, padx = (25, 0), pady = 5 )

      label = Label( text = 'Status :     No process is running', font='Helvetica 10 bold' )
      label.grid( row = 2, column = 0, padx = (25, 0), pady = 5 )

      button2 = Button( window, text = "  Send mails   ", command = create_thread )
      button2.grid( row = 2, column = 1, padx = (25, 0), pady = 10 )

      mainloop()

if __name__ == '__main__':
      main()
