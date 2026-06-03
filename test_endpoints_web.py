# import urllib.request
# import urllib.error
# import json
# import sys
# import os

# BASE_URL = "http://localhost:5000"

# def make_request(path, method="GET", payload=None):
#     url = f"{BASE_URL}{path}"
#     req = urllib.request.Request(url, method=method)
    
#     if payload is not None:
#         data = json.dumps(payload).encode('utf-8')
#         req.add_header('Content-Type', 'application/json')
#     else:
#         data = None
        
#     try:
#         with urllib.request.urlopen(req, data=data) as response:
#             body = response.read()
#             return response.status, body
#     except urllib.error.HTTPError as e:
#         return e.code, e.read()
#     except Exception as e:
#         print(f"Connection error to {url}: {e}")
#         return 0, str(e).encode()

# print("=== STARTING WEB API ENDPOINT VERIFICATION ===")

# # Test 1: POST /import-dummy
# print("\n--- Testing POST /import-dummy ---")
# status, body = make_request("/import-dummy", method="POST")
# print(f"Status: {status}")
# res_json = json.loads(body.decode('utf-8'))
# print(f"Response: {res_json}")
# if status != 200 or res_json.get('status') != 'success':
#     print("FAIL: /import-dummy failed")
#     sys.exit(1)

# # Test 2: GET /api/students (Retrieve imported dummy students)
# print("\n--- Testing GET /api/students ---")
# status, body = make_request("/api/students")
# print(f"Status: {status}")
# students = json.loads(body.decode('utf-8'))
# print(f"Imported students count: {len(students)}")
# if status != 200 or len(students) != 10:
#     print(f"FAIL: Expected 10 students, got {len(students)}")
#     sys.exit(1)
# print(f"First student: {students[0]['nama']} (NIM: {students[0]['nim']}, status: {students[0]['status_hadir']})")

# # Test 3: POST /blast-qr
# print("\n--- Testing POST /blast-qr ---")
# status, body = make_request("/blast-qr", method="POST")
# print(f"Status: {status}")
# res_json = json.loads(body.decode('utf-8'))
# print(f"Response: {res_json}")
# if status != 200 or res_json.get('status') != 'success':
#     print("FAIL: /blast-qr failed")
#     sys.exit(1)

# # Verify QR codes are generated
# qr_files = os.listdir(os.path.join('assets', 'qrcodes'))
# print(f"Verified: Generated {len(qr_files)} QR code PNG files in assets/qrcodes/.")
# if len(qr_files) != 10:
#     print("FAIL: QR code generation count mismatch")
#     sys.exit(1)

# # Test 4: POST /api/scan (Trigger check-in for first student)
# student_id = students[0]['id_mhs']
# student_name = students[0]['nama']
# print(f"\n--- Testing POST /api/scan for student {student_name} ({student_id}) ---")
# status, body = make_request("/api/scan", method="POST", payload={"id_mhs": student_id})
# print(f"Status: {status}")
# res_json = json.loads(body.decode('utf-8'))
# print(f"Response: {res_json}")
# if status != 200 or res_json.get('status') != 'success':
#     print("FAIL: /api/scan failed")
#     sys.exit(1)

# # Test 5: POST /api/scan (Trigger check-in again, should warning 'already scanned')
# print(f"\n--- Testing duplicate POST /api/scan for student {student_name} ---")
# status, body = make_request("/api/scan", method="POST", payload={"id_mhs": student_id})
# print(f"Status: {status}")
# res_json = json.loads(body.decode('utf-8'))
# print(f"Response: {res_json}")
# if status != 200 or res_json.get('status') != 'warning':
#     print("FAIL: Expected warning status for duplicate check-in")
#     sys.exit(1)

# # Test 5a: POST /api/students (Register a new student)
# print("\n--- Testing POST /api/students (New Registration) ---")
# new_student_payload = {
#     "nama": "Aditya Saputra",
#     "nim": "1202200999",
#     "email": "aditya@student.univ.ac.id"
# }
# status, body = make_request("/api/students", method="POST", payload=new_student_payload)
# print(f"Status: {status}")
# res_json = json.loads(body.decode('utf-8'))
# print(f"Response: {res_json}")
# if status != 201 or res_json.get('status') != 'success':
#     print("FAIL: /api/students registration failed")
#     sys.exit(1)
# new_student_id = res_json['student']['id_mhs']
# print(f"Verified: Registered student ID {new_student_id} and generated QR ticket.")

# # Test 6: GET /export-autocad (Export seating plot DXF)
# print("\n--- Testing GET /export-autocad ---")
# status, body = make_request("/export-autocad")
# print(f"Status: {status}")
# if status != 200 or len(body) < 1000: # Simple checks that we got valid DXF binary contents
#     print("FAIL: /export-autocad did not return a valid DXF file")
#     sys.exit(1)
# print(f"Verified: Downloaded DXF seating plot file successfully (size: {len(body)} bytes).")

# print("\n=== ALL WEB API VERIFICATION TESTS PASSED SUCCESSFULLY ===")
