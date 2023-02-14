from django.db import models

# Create your models here.

class ExcelFile(models.Model):
    file = models.FileField(upload_to='excel_files')
    uploaded_at = models.DateTimeField(auto_now_add=True)