from django.test import TestCase
from rest_framework.test import APIClient
from django.urls import reverse
from .models import Registration, Attendance
from django.utils import timezone


class AttendanceAPITestCase(TestCase):
	def setUp(self):
		self.client = APIClient()
		# create a registration to reference
		self.reg = Registration.objects.create(lrn='TESTLRN001', student='Test Student', sex='Male')

	def test_post_record_attendance(self):
		# POST to /api/attendance/ should create an Attendance
		res = self.client.post('/api/attendance/', {'lrn': self.reg.lrn}, format='json')
		self.assertIn(res.status_code, (201,))
		self.assertEqual(Attendance.objects.count(), 1)
		a = Attendance.objects.first()
		self.assertEqual(a.student.lrn, self.reg.lrn)

	def test_get_attendances_list(self):
		# create an attendance record first
		Attendance.objects.create(student=self.reg, time=timezone.now())
		# GET the list from the viewset endpoint
		res = self.client.get('/api/attendances/')
		self.assertEqual(res.status_code, 200)
		data = res.json()
		self.assertTrue(isinstance(data, list))
		self.assertGreaterEqual(len(data), 1)

	def test_get_attendance_fallback_endpoint(self):
		# ensure GET /api/attendance/ also returns data (fallback compatibility)
		Attendance.objects.create(student=self.reg, time=timezone.now())
		res = self.client.get('/api/attendance/')
		self.assertEqual(res.status_code, 200)
		data = res.json()
		self.assertTrue(isinstance(data, list))
