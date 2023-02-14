import openpyxl
from django.shortcuts import render, redirect
from .forms import UploadFileForm
from .models import ExcelFile
from twilio.rest import Client
from twilio.base.exceptions import TwilioRestException

def upload_file(request):
    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            # Save the uploaded file to the database
            file = ExcelFile(file=request.FILES['file'])
            file.save()

            # Load the Excel file using openpyxl
            wb = openpyxl.load_workbook(file.file.path)
            sheet = wb.active
            # Looping for mobile number
            for row in sheet.iter_rows(values_only=True):
                # row[0] should contain the mobile number
                number = row[0]
                
                # Twilio account SID and auth token
                account_sid = 'AC4cb9cd6934d4e5fcdbe9531af18776cd'
                auth_token = '1926e4210f1867650da220ea465189cf'
                client = Client(account_sid, auth_token)

                # Twilio WhatsApp number
                from_whatsapp_number = 'whatsapp:+15075785819'
                if number:
                    to_whatsapp_number = f'whatsapp:+91{number}'
                
                try:
                    message = client.messages.create(
                    body='Excel Testing',
                    from_=from_whatsapp_number,
                    to=to_whatsapp_number,
                    )
                except TwilioRestException as e:
                    print("Error: " + str(e))

            return redirect('success')
    else:
        form = UploadFileForm()
    return render(request, 'MyApp.html', {'form': form})

def success(request):
    return render(request, 'success.html')