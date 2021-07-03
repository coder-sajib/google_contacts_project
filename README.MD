' ZOOM
' https://us02web.zoom.us/j/7126833038

' PURPOSE
' Correct the contacts.csv export from contacts.google.com so that they all include the country code resulting in at least 11 digits for each phone number.
' From that cleaned file, create a list of names and numbers in output.txt based on the applied filter being carefull not include rows that have been filtered out.
' Likewise, create a list of names and email addresses below the names and numbers in the same output.txt.
' The only name column to use is the one called "name" and data in the other name columns can be removed.

' TESTING
' To test this code, simply download contacts.csv from contacts.google.com and run the macro against active sheet.

' AVAILABILITY
' My working hours are between 10 AM - Midnight Central Time US
' My current time is accurately reflected at https://time.is/CT

' CHANGES
' Added msgbox after completion of execution
' Code validates all the  phone fields
' Handled ":::"  exception , phone no will be splitted to next phone value cell 
' Compact code
' Code removes "("   "("   "-," & spaces In phone no1
' If  text & mobile no are  in same cell, text will be moved to next column
' Code filters all phone  no columns
' If phone value has ::: it will spilt & move the 2nd number to next phone value
' User will be notified once test done
' Output.txt is saved as text with format
' Name, phone no1
' Name, phone no2
' Name, phone no1
' added new pattern
