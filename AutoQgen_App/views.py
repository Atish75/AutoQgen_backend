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