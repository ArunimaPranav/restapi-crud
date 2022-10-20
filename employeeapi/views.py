from csv import excel
from datetime import datetime
from fileinput import filename
from http.client import HTTPResponse
from tkinter import font
from urllib import response
from rest_framework.response import Response
from django.shortcuts import get_object_or_404
from rest_framework import viewsets,status,filters
from . models import Employee
from .serializers import EmployeeSerializer
from employeeapi import serializers
import xlwt
import xlrd


# here is the views 

class EmployeeViewset(viewsets.ModelViewSet):
    queryset = Employee.objects.all()
    filter_backends = (filters.SearchFilter,)
    search_fields = ['fullname', 'emp_code']
    serializer_class = serializers.EmployeeSerializer
    
    
    def create(self, request) :
        serializer = EmployeeSerializer(data=request.data)
        print(request.data)
        if serializer.is_valid() :
            serializer.save()
            return Response(serializer.data, status=status.HTTP_201_CREATED)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
    
    def update(self, request, pk) :
        try :
            queryset = Employee.objects.all()
            employee = get_object_or_404(queryset, pk=pk)
            serializer = EmployeeSerializer(instance=employee, data=request.data)
            if serializer.is_valid():
                serializer.save()
                return Response(serializer.data, status=status.HTTP_200_OK)
            return Response(status=status.HTTP_400_BAD_REQUEST)
        except Exception:
            return Response(status=status.HTTP_400_BAD_REQUEST)
    
    
 
    
    def destroy(self, request, pk) :
        try :
            queryset = Employee.objects.all()
            employee = get_object_or_404(queryset, pk=pk)
            employee.delete()
            return Response(status=status.HTTP_204_NO_CONTENT)
        except Exception:
            return Response(status=status.HTTP_400_BAD_REQUEST)
   
    def retrieve(self,request,*args,**kwargs):
        params = kwargs
        print(params['fullname'])
        employees = Employee.objects.filter(fullname = params['fullname'])
        serializer = EmployeeSerializer(employees,many=True)
        return Response(serializer.data)

def export_excel(request):
    response = HTTPResponse(content_type = 'application/ms-excel')
    response['Content-Deposition'] = 'attachment;filename=Employees'+ \
    str(datetime.datetime.now())+'.xls'
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet ('Employees')
    row_num = 0
    font_style = xlwt.XFStyle()
    font_style.font.bold = True
    columns = ['emp_code','fullname','mobile']
    for col_num in range(len(columns)):
        ws.write(row_num,col_num,columns[col_num],font_style)
    font.font.bold = xlwt.XFStyle()
    rows = Employee.objects.filter(owner = request.user).values_list('emp_code','fullname','mobile')
    for row in rows:
        row_num +=1
        for col_num in range (len(row)):
            ws.write(row_num,col_num,str(row[col_num]),font_style)
    wb.save(response)
    return response
    
