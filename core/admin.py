# core/admin.py
from django.contrib import admin
from .models import FileCategory, File, FileAccessLog, Ticket#, TicketComment


@admin.register(FileCategory)
class FileCategoryAdmin(admin.ModelAdmin):
    list_display = ('id', 'name', 'file_count')
    def file_count(self, obj):
        return obj.file_set.filter(is_deleted=False).count()
    file_count.short_description = 'Number of Files'

@admin.register(File)
class FileAdmin(admin.ModelAdmin):
    list_display = ('title', 'category', 'uploaded_by', 'upload_date', 'access_level', 'is_deleted')
    list_filter = ('category', 'access_level', 'is_deleted')
    search_fields = ('title',)

@admin.register(FileAccessLog)
class FileAccessLogAdmin(admin.ModelAdmin):
    list_display = ('file', 'accessed_by', 'access_time')
    list_filter = ('access_time',)

@admin.register(Ticket)
class TicketAdmin(admin.ModelAdmin):
    list_display = ('title', 'created_by', 'assigned_to', 'status', 'priority', 'created_at')
    list_filter = ('status', 'priority')
    search_fields = ('title', 'description')

"""@admin.register(TicketComment)
class TicketCommentAdmin(admin.ModelAdmin):
    list_display = ('ticket', 'commented_by', 'created_at')
    list_filter = ('created_at',)"""
