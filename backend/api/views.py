from rest_framework import viewsets, status
from rest_framework.decorators import api_view, action
from rest_framework.response import Response
from django.utils import timezone
from django.conf import settings
import os
from .models import Registration, Attendance, DroppedRegistration
from .serializers import RegistrationSerializer, AttendanceSerializer, DroppedRegistrationSerializer
from rest_framework.decorators import permission_classes
from rest_framework.permissions import AllowAny
import io
try:
    import openpyxl
    from openpyxl.styles import Font, Alignment
except Exception:
    openpyxl = None

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


@api_view(['DELETE'])
def clear_attendance(request):
    """Clear attendance records. Query param 'scope' accepts 'today' or 'all'. Defaults to 'today'."""
    scope = request.GET.get('scope', 'today')
    try:
        if scope == 'all':
            Attendance.objects.all().delete()
            return Response({"message": "All attendance records deleted."}, status=status.HTTP_204_NO_CONTENT)
        else:
            today = timezone.localdate()
            Attendance.objects.filter(time__date=today).delete()
            return Response({"message": "Today's attendance deleted."}, status=status.HTTP_204_NO_CONTENT)
    except Exception as e:
        return Response({'error': str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


class AttendanceViewSet(viewsets.ModelViewSet):
    queryset = Attendance.objects.all()
    serializer_class = AttendanceSerializer

    @action(detail=False, methods=['get'])
    def today(self, request):
        today_date = now().date()
        attendances = Attendance.objects.filter(timestamp__date=today_date)
        serializer = self.get_serializer(attendances, many=True)
        return Response(serializer.data)


@api_view(['POST'])
def upload_excel(request):
    """Accept an uploaded Excel file and save it to the backend 'exports' folder."""
    f = request.FILES.get('file')
    if not f:
        return Response({'error': 'No file uploaded'}, status=status.HTTP_400_BAD_REQUEST)

    # Ensure exports directory exists under BASE_DIR
    out_dir = settings.BASE_DIR / 'exports'
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception:
        # In some environments BASE_DIR may be a string
        out_dir = os.path.join(str(settings.BASE_DIR), 'exports')
        os.makedirs(out_dir, exist_ok=True)

    filename = f"{timezone.now().strftime('%Y%m%d_%H%M%S')}_{f.name}"
    dest_path = out_dir / filename if hasattr(out_dir, '__truediv__') else os.path.join(out_dir, filename)

    try:
        # Write uploaded chunks to disk
        with open(dest_path, 'wb') as dest:
            for chunk in f.chunks():
                dest.write(chunk)
    except Exception as e:
        return Response({'error': f'Failed to save file: {e}'}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

    # Provide a URL so clients can download the saved file. During development we expose
    # the 'exports' folder via a static serve URL defined in backend/urls.py
    download_url = f"/exports/{filename}"
    return Response({'message': 'Saved', 'filename': filename, 'url': download_url}, status=status.HTTP_201_CREATED)


@api_view(['GET'])
def generate_excel_export(request):
    """Generate an Excel export server-side and save it to exports/. Returns download URL."""
    if openpyxl is None:
        return Response({'error': 'openpyxl not installed on server'}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

    regs = Registration.objects.all().order_by('student')
    attends = Attendance.objects.all()

    wb = openpyxl.Workbook()
    # create a single sheet for current month
    today = timezone.localdate()
    month_name = today.strftime('%b').upper()
    ws = wb.active
    ws.title = month_name

    # Header row: Day numbers 1..31 starting at column 3
    ws.cell(row=1, column=1, value='Student')
    for d in range(1, 32):
        ws.cell(row=1, column=2 + d, value=d)

    # Build attendance map keyed by lrn or student name lowercased
    att_map = {}
    for a in attends:
        try:
            d = a.time.date()
        except Exception:
            continue
        if d.month != today.month or d.year != today.year:
            continue
        key = (a.student.lrn or a.student.student).strip().lower()
        att_map.setdefault(key, set()).add(d.day)

    row = 2
    for r in regs:
        name = r.student or r.lrn
        key = (r.lrn or r.student).strip().lower()
        ws.cell(row=row, column=1, value=name)
        for d in range(1, today.day + 1):
            cell = ws.cell(row=row, column=2 + d)
            if key in att_map and d in att_map[key]:
                cell.value = 'P'
                cell.font = Font(bold=True, color='FF00A050')
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.value = 'X'
                cell.font = Font(color='FF000000')
                cell.alignment = Alignment(horizontal='center', vertical='center')
        row += 1

    # Save workbook to exports
    out_dir = settings.BASE_DIR / 'exports'
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception:
        out_dir = os.path.join(str(settings.BASE_DIR), 'exports')
        os.makedirs(out_dir, exist_ok=True)

    filename = f"attendance_server_export_{timezone.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    dest_path = out_dir / filename if hasattr(out_dir, '__truediv__') else os.path.join(out_dir, filename)
    wb.save(dest_path)

    download_url = f"/exports/{filename}"
    return Response({'filename': filename, 'url': download_url}, status=status.HTTP_201_CREATED)


@api_view(['GET'])
@permission_classes([AllowAny])
def registrations_grouped(request):
    """Return registrations grouped by sex (Male/Female) and sorted alphabetically by student name."""
    regs = Registration.objects.all()
    males = regs.filter(sex__iexact='male').order_by('student')
    females = regs.filter(sex__iexact='female').order_by('student')

    male_ser = RegistrationSerializer(males, many=True)
    female_ser = RegistrationSerializer(females, many=True)

    return Response({
        'male': male_ser.data,
        'female': female_ser.data,
    })


@api_view(['POST'])
def drop_registration(request, pk):
    """Copy a registration into DroppedRegistration then delete the original.

    This preserves a server-side record of the dropped student so it can be
    restored later without relying on client localStorage.
    """
    try:
        reg = Registration.objects.get(pk=pk)
    except Registration.DoesNotExist:
        return Response({'error': 'Registration not found'}, status=status.HTTP_404_NOT_FOUND)

    try:
        # create dropped copy
        dropped = DroppedRegistration.objects.create(
            lrn=reg.lrn,
            student=reg.student,
            sex=reg.sex,
            parent=reg.parent,
            guardian=reg.guardian,
            original_id=reg.id,
        )

        # delete original registration (this will cascade-delete attendances)
        reg.delete()

        ser = DroppedRegistrationSerializer(dropped)
        return Response(ser.data, status=status.HTTP_201_CREATED)
    except Exception as e:
        # Return a helpful message for debugging (dev only)
        return Response({'error': f'Failed to drop registration: {e}'}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


@api_view(['GET'])
def dropped_list(request):
    try:
        drops = DroppedRegistration.objects.all().order_by('-dropped_at')
        ser = DroppedRegistrationSerializer(drops, many=True)
        return Response(ser.data)
    except Exception as e:
        return Response({'error': f'Failed to fetch dropped list: {e}'}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


@api_view(['POST'])
def restore_dropped(request, pk):
    """Restore a dropped registration back into the main Registration table.

    This will create a new Registration using the stored data and remove the
    DroppedRegistration record.
    """
    try:
        dropped = DroppedRegistration.objects.get(pk=pk)
    except DroppedRegistration.DoesNotExist:
        return Response({'error': 'Dropped registration not found'}, status=status.HTTP_404_NOT_FOUND)

    try:
        # Prevent duplicate LRN
        if Registration.objects.filter(lrn=dropped.lrn).exists():
            return Response({'error': 'Registration with same LRN already exists'}, status=status.HTTP_400_BAD_REQUEST)

        reg = Registration.objects.create(
            lrn=dropped.lrn,
            student=dropped.student,
            sex=dropped.sex,
            parent=dropped.parent,
            guardian=dropped.guardian,
        )

        # delete dropped record
        dropped.delete()

        ser = RegistrationSerializer(reg)
        return Response(ser.data, status=status.HTTP_201_CREATED)
    except Exception as e:
        return Response({'error': f'Failed to restore dropped registration: {e}'}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


@api_view(['DELETE'])
def delete_dropped(request, pk):
    try:
        dropped = DroppedRegistration.objects.get(pk=pk)
    except DroppedRegistration.DoesNotExist:
        return Response({'error': 'Dropped registration not found'}, status=status.HTTP_404_NOT_FOUND)
    try:
        dropped.delete()
        return Response({'message': 'Dropped registration deleted'}, status=status.HTTP_204_NO_CONTENT)
    except Exception as e:
        return Response({'error': f'Failed to delete dropped registration: {e}'}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)