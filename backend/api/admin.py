from django.contrib import admin
from .models import Registration, Attendance

@admin.register(Registration)
class RegistrationAdmin(admin.ModelAdmin):
    list_display = ('student', 'lrn', 'sex', 'created_at')
    search_fields = ('student', 'lrn', 'parent', 'guardian')

@admin.register(Attendance)
class AttendanceAdmin(admin.ModelAdmin):
    list_display = ('student', 'time')
    list_filter = ('time',)
    search_fields = ('student__student', 'student__lrn')
