from django.db import models
import os

class UploadedFile(models.Model):
    FILE_TYPE_CHOICES = [
        ('Préparation_PL', 'Fichier Préparation PL'),
        ('zzzz', 'Fichier zzzz'),
    ]
    
    file = models.FileField(upload_to='uploads/')
    file_type = models.CharField(max_length=20, choices=FILE_TYPE_CHOICES)
    uploaded_at = models.DateTimeField(auto_now_add=True)
    original_name = models.CharField(max_length=255)

    def __str__(self):
        return f"{self.file_type} - {self.original_name}"

class GeneratedFile(models.Model):
    FILE_TYPE_CHOICES = [
        ('excel', 'Fichier Excel'),
        ('pdf', 'Fichier PDF'),
    ]
    
    file = models.FileField(upload_to='generated/')
    file_type = models.CharField(max_length=10, choices=FILE_TYPE_CHOICES)
    container_name = models.CharField(max_length=100)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.container_name} - {self.file_type}"
    
    def filename(self):
        return os.path.basename(self.file.name)