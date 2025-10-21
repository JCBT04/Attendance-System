from rest_framework import viewsets, status
from rest_framework.decorators import api_view, action
from rest_framework.response import Response
from django.utils import timezone
from .models import Registration, Attendance
from .serializers import RegistrationSerializer, AttendanceSerializer

class RegistrationViewSet(viewsets.ModelViewSet):
    queryset = Registration.objects.all().order_by('student')
    serializer_class = RegistrationSerializer


@api_view(['POST'])
def record_attendance(request):
    """Record attendance via QR scan"""
    data = request.data
    lrn = data.get('lrn')
    if not lrn:
        return Response({'error': 'LRN is required'}, status=status.HTTP_400_BAD_REQUEST)
    
    try:
        student = Registration.objects.get(lrn=lrn)
    except Registration.DoesNotExist:
        return Response({'error': 'Student not found'}, status=status.HTTP_404_NOT_FOUND)
    
    Attendance.objects.create(student=student, time=timezone.now())
    return Response({'message': f'Attendance recorded for {student.student}'}, status=status.HTTP_201_CREATED)


@api_view(['GET'])
def attendance_today(request):
    """Fetch today’s attendance"""
    today = timezone.localdate()
    records = Attendance.objects.filter(time__date=today)
    serializer = AttendanceSerializer(records, many=True)
    return Response(serializer.data)

@api_view(['DELETE'])
def clear_all_registrations(request):
    Registration.objects.all().delete()
    return Response({"message": "All registrations deleted."}, status=status.HTTP_204_NO_CONTENT)


class AttendanceViewSet(viewsets.ModelViewSet):
    queryset = Attendance.objects.all()
    serializer_class = AttendanceSerializer

    @action(detail=False, methods=['get'])
    def today(self, request):
        today_date = now().date()
        attendances = Attendance.objects.filter(timestamp__date=today_date)
        serializer = self.get_serializer(attendances, many=True)
        return Response(serializer.data)