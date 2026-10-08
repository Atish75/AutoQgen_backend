from django.db import models

# Create your models here.

class ProcessedDocument(models.Model):
    title = models.CharField(max_length=255)
    original_filename = models.CharField(max_length=255)
    cloudinary_excel_url = models.URLField(max_length=1000)
    cloudinary_public_id = models.CharField(max_length=255)
    total_records = models.IntegerField(default=0)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    def __str__(self):
        return f"{self.title} (ID: {self.id})"