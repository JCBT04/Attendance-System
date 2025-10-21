from django.urls import path, include
from rest_framework.routers import DefaultRouter
from .views import RegistrationViewSet, record_attendance, attendance_today, clear_all_registrations

router = DefaultRouter()
router.register(r'registrations', RegistrationViewSet, basename='registration')

urlpatterns = [
    # ✅ Place clear_all FIRST so it isn't mistaken for a registration ID
    path('registrations/clear_all/', clear_all_registrations),
    path('attendance/', record_attendance),
    path('attendance/today/', attendance_today),
    path('', include(router.urls)),  # router goes last
]