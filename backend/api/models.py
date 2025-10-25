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


class DroppedRegistration(models.Model):
    """Server-side store for dropped registrations so drops persist across clients.

    We copy the essential fields from Registration here when a student is dropped.
    """
    lrn = models.CharField(max_length=20)
    student = models.CharField(max_length=100)
    sex = models.CharField(max_length=10)
    parent = models.CharField(max_length=100, blank=True)
    guardian = models.CharField(max_length=100, blank=True)
    dropped_at = models.DateTimeField(auto_now_add=True)
    original_id = models.IntegerField(null=True, blank=True)

    def __str__(self):
        return f"Dropped: {self.student} ({self.lrn})"
