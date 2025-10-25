from rest_framework import serializers
from .models import Registration, Attendance, DroppedRegistration

class RegistrationSerializer(serializers.ModelSerializer):
    class Meta:
        model = Registration
        fields = '__all__'


class AttendanceSerializer(serializers.ModelSerializer):
    student_name = serializers.CharField(source='student.student', read_only=True)

    class Meta:
        model = Attendance
        fields = ['id', 'student', 'student_name', 'time']


class DroppedRegistrationSerializer(serializers.ModelSerializer):
    class Meta:
        model = DroppedRegistration
        fields = '__all__'
