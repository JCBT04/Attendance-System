from django.urls import path, include
from rest_framework.routers import DefaultRouter
from .views import (
    RegistrationViewSet,
    record_attendance,
    attendance_today,
    clear_all_registrations,
    clear_attendance,
    upload_excel,
    generate_excel_export,
    AttendanceViewSet,
)
from .views import registrations_grouped

router = DefaultRouter()
router.register(r'registrations', RegistrationViewSet, basename='registration')
router.register(r'attendances', AttendanceViewSet, basename='attendance')

urlpatterns = [
    # ✅ Place clear_all FIRST so it isn't mistaken for a registration ID
    path('registrations/clear_all/', clear_all_registrations),
    path('registrations/grouped/', registrations_grouped),
    path('attendance/clear/', clear_attendance),
    path('attendance/generate_excel/', generate_excel_export),
    path('attendance/', record_attendance),
    path('attendance/today/', attendance_today),
    path('attendance/upload_excel/', upload_excel),
    path('', include(router.urls)),  # router goes last
]