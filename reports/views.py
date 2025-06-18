import os
import json
from django.shortcuts import render
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from .parser_wrapper import process_credit_report



@csrf_exempt
def get_report(request):
    print("🔥 get_report view called")

    if request.method == 'POST':
        print("📩 POST request received")

        if request.FILES.get('pdf'):
            print("📎 PDF file received")

            uploaded_pdf = request.FILES['pdf']
            temp_path = f'temp/{uploaded_pdf.name}'

            os.makedirs('temp', exist_ok=True)
            with open(temp_path, 'wb') as f:
                for chunk in uploaded_pdf.chunks():
                    f.write(chunk)

            print("💾 File saved at:", temp_path)

            try:
                data = process_credit_report(temp_path)
                print("✅ Data returned:", data)
                print("🧪 Type of returned data:", type(data))
                return JsonResponse({'status': 'success', 'data': data})
            except Exception as e:
                print("❌ Exception:", e)
                return JsonResponse({'status': 'error', 'message': str(e)})

        else:
            print("🚫 PDF not found in request")

    else:
        print("🚫 Not a POST request")

    return JsonResponse({'status': 'error', 'message': 'Invalid request'})

def index(request):
    return render(request, 'reports/MyFreeScoreNow.html')
