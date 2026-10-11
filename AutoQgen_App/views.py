# AutoQgen_App/views.py

import uuid
import cloudinary
import pandas as pd
from django.shortcuts import render, get_object_or_404
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from rest_framework.parsers import MultiPartParser, FormParser
from rest_framework import serializers

from .models import ProcessedDocument
from .serializers import ProcessedDocumentSerializer
from .services import process_docx_file, upload_dataframe_as_excel


class FileUploadSerializer(serializers.Serializer):
    file = serializers.FileField()
    title = serializers.CharField(max_length=255, required=False)


class DocxUploadView(APIView):
    parser_classes = (MultiPartParser, FormParser)
    serializer_class = FileUploadSerializer

    def get_serializer(self, *args, **kwargs):
        return FileUploadSerializer(*args, **kwargs)

    def get(self, request):
        return Response({"message": "Upload a .docx file using POST request below."})

    def post(self, request):
        serializer = FileUploadSerializer(data=request.data)
        if not serializer.is_valid():
            return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)

        file_obj = serializer.validated_data['file']
        if not file_obj.name.endswith('.docx'):
            return Response(
                {"error": "Please upload a valid .docx file."}, 
                status=status.HTTP_400_BAD_REQUEST
            )

        title = serializer.validated_data.get('title', file_obj.name.rsplit('.', 1)[0])
        public_id = f"doc_{uuid.uuid4().hex[:8]}"

        df = process_docx_file(file_obj)

        if df.empty:
            return Response(
                {"error": "No content found in document."}, 
                status=status.HTTP_400_BAD_REQUEST
            )

        excel_url, c_public_id = upload_dataframe_as_excel(df, public_id)

        doc_record = ProcessedDocument.objects.create(
            title=title,
            original_filename=file_obj.name,
            cloudinary_excel_url=excel_url,
            cloudinary_public_id=c_public_id,
            total_records=len(df)
        )

        return Response({
            "message": "File processed & uploaded to Cloudinary successfully!",
            "data": ProcessedDocumentSerializer(doc_record).data
        }, status=status.HTTP_201_CREATED)


class DocumentDataView(APIView):
    """
    GET: Cloudinary Excel read karke JSON deta hai.
    PUT: Edited data se Excel update karke Cloudinary overwrite karta hai.
    DELETE: Cloudinary se file aur Database se record delete karta hai.
    """
    def get(self, request, pk):
        try:
            doc = ProcessedDocument.objects.get(pk=pk)
            # Cloudinary URL se fresh Excel read
            df = pd.read_excel(doc.cloudinary_excel_url)
            records = df.fillna("").to_dict(orient="records")
            return Response({
                "document_id": doc.id,
                "title": doc.title,
                "cloudinary_url": doc.cloudinary_excel_url,
                "records": records
            }, status=status.HTTP_200_OK)
        except ProcessedDocument.DoesNotExist:
            return Response({"error": "Document not found."}, status=status.HTTP_404_NOT_FOUND)

    def put(self, request, pk):
        try:
            doc = ProcessedDocument.objects.get(pk=pk)
            updated_records = request.data.get("records")

            if not isinstance(updated_records, list):
                return Response(
                    {"error": "Expected 'records' to be a list of rows."}, 
                    status=status.HTTP_400_BAD_REQUEST
                )

            # Updated JSON se DataFrame banayein
            df_updated = pd.DataFrame(updated_records)

            # Sabhi mandatory columns ensure karein aur string cast karein
            expected_cols = ["item_id", "type", "content", "equations", "diagram_urls"]
            for col in expected_cols:
                if col not in df_updated.columns:
                    df_updated[col] = ""
                else:
                    df_updated[col] = df_updated[col].fillna("").astype(str).replace({'nan': '', 'None': ''})

            # Cloudinary par overwrite karein
            clean_public_id = doc.cloudinary_public_id.replace('.xlsx', '')
            excel_url, _ = upload_dataframe_as_excel(df_updated, clean_public_id)

            # FIX: Database me naya Cloudinary URL aur records update karein
            doc.cloudinary_excel_url = excel_url
            doc.total_records = len(df_updated)
            doc.save()

            return Response({
                "message": "File updated and overwritten on Cloudinary successfully!",
                "cloudinary_url": excel_url,
                "total_records": doc.total_records
            }, status=status.HTTP_200_OK)

        except ProcessedDocument.DoesNotExist:
            return Response({"error": "Document not found."}, status=status.HTTP_404_NOT_FOUND)
        except Exception as e:
            return Response({"error": f"Failed to update: {str(e)}"}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

    def delete(self, request, pk):
        try:
            doc = ProcessedDocument.objects.get(pk=pk)

            cloudinary.uploader.destroy(
                doc.cloudinary_public_id,
                resource_type="raw",
                invalidate=True
            )

            doc_title = doc.title
            doc.delete()

            return Response({
                "message": f"'{doc_title}' and its Cloudinary file deleted successfully!"
            }, status=status.HTTP_200_OK)

        except ProcessedDocument.DoesNotExist:
            return Response({"error": "Document not found."}, status=status.HTTP_404_NOT_FOUND)
        except Exception as e:
            return Response({"error": f"Failed to delete: {str(e)}"}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)


def dashboard_view(request):
    """Saari uploaded files list karta hai."""
    documents = ProcessedDocument.objects.all().order_by('-created_at')
    return render(request, 'dashboard.html', {'documents': documents})


def editor_view(request, pk):
    """Specific document ko edit karne ka view."""
    document = get_object_or_404(ProcessedDocument, pk=pk)
    return render(request, 'editor.html', {'document': document})

# AutoQgen_App/views.py me add karein
from django.http import HttpResponse
from .services import generate_exam_docx

def exam_paper_view(request, pk):
    """Formal HTML exam template for printing / saving as PDF."""
    doc = get_object_or_404(ProcessedDocument, pk=pk)
    df = pd.read_excel(doc.cloudinary_excel_url)
    records = df.fillna("").to_dict(orient="records")

    # Split diagram URLs for template iteration
    for r in records:
        urls = r.get("diagram_urls", "")
        r["images_list"] = [u.strip() for u in urls.split("|") if u.strip().startswith("http")]

    return render(request, 'exam_template.html', {
        'document': doc,
        'records': records
    })

def export_docx_view(request, pk):
    """Generated Question paper as .docx file download."""
    doc = get_object_or_404(ProcessedDocument, pk=pk)
    df = pd.read_excel(doc.cloudinary_excel_url)
    records = df.fillna("").to_dict(orient="records")

    docx_stream = generate_exam_docx(doc.title, records)
    response = HttpResponse(
        docx_stream.getvalue(),
        content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    )
    response['Content-Disposition'] = f'attachment; filename="{doc.title}_question_paper.docx"'
    return response

class ImageDirectUploadView(APIView):
    """
    Editor UI se direct local image Cloudinary par upload karta hai
    aur naya secure URL return karta hai.
    """
    parser_classes = (MultiPartParser, FormParser)

    def post(self, request):
        image_file = request.FILES.get('image')
        if not image_file:
            return Response({"error": "No image file provided."}, status=status.HTTP_400_BAD_REQUEST)

        try:
            upload_res = cloudinary.uploader.upload(
                image_file,
                folder="exam_diagrams",
                resource_type="image"
            )
            return Response({
                "image_url": upload_res.get("secure_url")
            }, status=status.HTTP_200_OK)
        except Exception as e:
            return Response({"error": str(e)}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)

import re

class ImageDirectDeleteView(APIView):
    """
    Cloudinary image URL receive karta hai aur Cloudinary storage se
    image permanently delete karta hai.
    """
    def post(self, request):
        image_url = request.data.get('image_url', '').strip()
        if not image_url:
            return Response({"error": "No image URL provided."}, status=status.HTTP_400_BAD_REQUEST)

        try:
            # URL se public_id extract karein:
            # Example: https://res.cloudinary.com/.../image/upload/v12345/exam_diagrams/abc123xyz.png
            # -> Target Public ID: "exam_diagrams/abc123xyz"
            match = re.search(r'/upload/(?:v\d+/)?(.+?)(?:\.[a-zA-Z0-9]+)?$', image_url)
            if not match:
                return Response({"error": "Invalid Cloudinary URL format."}, status=status.HTTP_400_BAD_REQUEST)

            public_id = match.group(1)

            # Cloudinary se delete karein
            destroy_res = cloudinary.uploader.destroy(
                public_id,
                resource_type="image",
                invalidate=True
            )

            return Response({
                "message": "Image deleted from Cloudinary successfully!",
                "result": destroy_res.get("result")
            }, status=status.HTTP_200_OK)

        except Exception as e:
            return Response({"error": f"Failed to delete image: {str(e)}"}, status=status.HTTP_500_INTERNAL_SERVER_ERROR)