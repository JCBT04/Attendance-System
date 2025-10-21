from django.db import models

class Registration(models.Model):
    lrn = models.CharField(max_length=20, unique=True)
    student = models.CharField(max_length=100)
    sex = models.CharField(max_length=10)
    parent = models.CharField(max_length=100)
    guardian = models.CharField(max_length=100)
    created_at = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.student} ({self.lrn})"


class Attendance(models.Model):
    student = models.ForeignKey(Registration, on_delete=models.CASCADE)
    time = models.DateTimeField(auto_now_add=True)

    def __str__(self):
        return f"{self.student.student} - {self.time.strftime('%Y-%m-%d %H:%M:%S')}"
