# WhatsappBot
This is the code for a whatsapp chatbot to send payment reminders to customers.   
Uses the Whatsapp Buisness API and the openpyxl library.   
It retrieves customer records from an Excel file and traverses it to send the messages.   
Calculates the total amount paid by each customer over multiple different sheets and subtracts that from the total amount due to find out the amount owed.   
Generates log of status of message(sent,unsent,undervalued) in a seperate excel file for user to review.   
Was tailored to fit a specific excel file but can be adapted to different files.   

