from django.shortcuts import render

from django.http import HttpResponse





def search(request):
    return HttpResponse("My seacrh!")
    
def index(request):
    return HttpResponse("My index!")