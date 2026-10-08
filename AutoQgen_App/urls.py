from django.urls import path
from .views import DocxUploadView,DocumentDataView,dashboard_view,editor_view

urlpatterns = [
    path('', dashboard_view, name='dashboard'),
    path('edit/<int:pk>/', editor_view, name='editor'),
    path('upload/', DocxUploadView.as_view(), name='docx-upload'),
    path('document/<int:pk>/', DocumentDataView.as_view(), name='document-detail-edit'),
]