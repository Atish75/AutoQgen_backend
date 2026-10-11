from django.urls import path
from .views import (DocxUploadView,DocumentDataView,dashboard_view,editor_view,exam_paper_view,export_docx_view,
                    ImageDirectUploadView,ImageDirectDeleteView )

urlpatterns = [
    path('', dashboard_view, name='dashboard'),
    path('edit/<int:pk>/', editor_view, name='editor'),
    path('upload/', DocxUploadView.as_view(), name='docx-upload'),
    path('document/<int:pk>/', DocumentDataView.as_view(), name='document-detail-edit'),
    path('paper/<int:pk>/', exam_paper_view, name='exam-paper-print'),
    path('export-docx/<int:pk>/', export_docx_view, name='export-docx'),
    path('upload-image/', ImageDirectUploadView.as_view(), name='image-direct-upload'),
    path('delete-image/', ImageDirectDeleteView.as_view(), name='image-direct-delete'),
]