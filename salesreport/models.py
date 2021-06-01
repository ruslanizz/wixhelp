from django.db import models

class SalesReport(models.Model):
    csv_from_wix=models.FileField(blank=True)
    xls_from_1c=models.FileField(blank=True)
    csv_ready_to_wix=models.FileField(blank=True)
    log=models.JSONField(blank=True)


