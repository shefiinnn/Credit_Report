import os
import json
from django.shortcuts import render
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from . import parser_wrapper

@csrf_exempt
def get_report(request):    
    if request.method == 'POST' and request.FILES.get('pdf'):
        uploaded_pdf = request.FILES['pdf']
        temp_path = f'temp/{uploaded_pdf.name}'
        
        os.makedirs('temp', exist_ok=True)
        with open(temp_path, 'wb') as f:
            for chunk in uploaded_pdf.chunks():
                f.write(chunk)

        try:
            data = parser_wrapper.process_credit_report(temp_path)
            return JsonResponse({'status': 'success', 'data': data})
        except Exception as e:
            return JsonResponse({'status': 'error', 'message': str(e)})

    return JsonResponse({'status': 'error', 'message': 'Invalid request'})

def index(request):
    return render(request, 'reports/MyFreeScoreNow.html')
