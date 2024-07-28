from django.db import models

# Create your models here.
class Employee(models.Model):
    Empid=models.CharField(max_length=20, unique=True)
    Name=models.CharField(max_length=100)
    CCDescpription= models.CharField(max_length=100)
    Personal_Area=models.CharField(max_length=100)
    Personal_Sub_Area=models.CharField(max_length=100)
    Designation=models.CharField(max_length=100)

    def __str__(self):
        return self.Name