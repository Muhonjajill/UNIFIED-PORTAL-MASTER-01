# forms.py
from django import forms
from .models import File, ProblemCategory, Ticket,TicketComment
from django.contrib.auth.models import User
from .models import Profile, Terminal, VersionControl

class LoginForm(forms.Form):
    username = forms.CharField()
    password = forms.CharField(widget=forms.PasswordInput)

class OTPForm(forms.Form):
    otp = forms.CharField(max_length=6)
    
class CustomUserCreationForm(forms.ModelForm):
    password = forms.CharField(label='Password', widget=forms.PasswordInput)
    password2 = forms.CharField(label='Confirm Password', widget=forms.PasswordInput)

    class Meta:
        model = User
        fields = ['first_name', 'last_name', 'username', 'email']

    def clean_password2(self):
        password = self.cleaned_data.get('password')
        password2 = self.cleaned_data.get('password2')

        if password and password2 and password != password2:
            raise forms.ValidationError("Passwords don’t match.")
        return password2

    def save(self, commit=True):
        user = super().save(commit=False)
        user.set_password(self.cleaned_data['password'])
        if commit:
            user.save()
        return user
    
class FileUploadForm(forms.ModelForm):
    class Meta:
        model = File
        fields = ['title', 'description', 'category', 'access_level', 'file']
        widgets = {
            'title': forms.TextInput(attrs={
                'class': 'form-control',
                'placeholder': 'Enter file title'
            }),
            'description': forms.Textarea(attrs={
                'class': 'form-control',
                'placeholder': 'Enter a brief description',
                'rows': 3
            }),
            'category': forms.Select(attrs={
                'class': 'form-control'
            }),
            'access_level': forms.Select(attrs={
                'class': 'form-control'
            }),
            'file': forms.ClearableFileInput(attrs={
                'class': 'form-control-file'
            }),
        }

class UserUpdateForm(forms.ModelForm):
    email = forms.EmailField()

    class Meta:
        model = User
        fields = ['username', 'email']

class ProfileUpdateForm(forms.ModelForm):
    class Meta:
        model = Profile
        fields = ['avatar']

class TicketForm(forms.ModelForm):
    class Meta:
        model = Ticket
        exclude = ['created_by', 'assigned_to', 'created_at', 'updated_at']
        widgets = {
            'title': forms.TextInput(attrs={'class': 'form-control'}),
            'description': forms.Textarea(attrs={'class': 'form-control', 'rows': 4}),
            'brts_unit': forms.Select(attrs={'class': 'form-control'}),
            'customer': forms.Select(attrs={'class':'form-control'}),
            'region': forms.Select(attrs={'class':'form-control'}),
            'problem_category': forms.Select(attrs={'class': 'form-control'}),
            'terminal': forms.Select(attrs={'class': 'form-control'}),
            'priority': forms.Select(attrs={'class': 'form-control'}),
            'status': forms.Select(attrs={'class': 'form-control'}),
            'responsible': forms.Select(attrs={'class': 'form-control'}),
        }



class ProblemCategoryForm(forms.ModelForm):
    class Meta:
        model = ProblemCategory
        fields = ['brts_unit', 'name']
        widgets = {
            'brts_unit': forms.Select(attrs={'class': 'form-control'}),
            'name': forms.TextInput(attrs={'class': 'form-control', 'placeholder': 'Enter category name'}),
        }

class TerminalForm(forms.ModelForm):
    class Meta:
        model = Terminal
        fields = ['customer', 'branch_name', 'cdm_name', 'serial_number', 'region', 'model', 'zone']
        widgets = {
            'customer': forms.Select(attrs={'class': 'form-control'}),
            'branch_name': forms.TextInput(attrs={'class': 'form-control'}),
            'cdm_name': forms.TextInput(attrs={'class': 'form-control'}),
            'serial_number': forms.TextInput(attrs={'class': 'form-control'}),
            'region': forms.Select(attrs={'class': 'form-control'}),
            'model': forms.TextInput(attrs={'class': 'form-control'}),
            'zone': forms.Select(attrs={'class': 'form-control'}),
        }
    
class TerminalUploadForm(forms.Form):
    file = forms.FileField(
        label='Upload CSV or Excel File',
        widget=forms.FileInput(attrs={'accept': '.csv, application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})
    )

class VersionControlForm(forms.ModelForm):
    class Meta:
        model = VersionControl
        fields = ['terminal', 'manufacturer', 'template', 'firmware', 'xfs', 'ejournal', 'responsible', 'app_version']
        widgets = {
            'terminal': forms.Select(attrs={'class': 'form-control'}),
            'manufacturer': forms.TextInput(attrs={'class': 'form-control'}),
            'template': forms.TextInput(attrs={'class': 'form-control'}),
            'firmware': forms.TextInput(attrs={'class': 'form-control'}),
            'xfs': forms.TextInput(attrs={'class': 'form-control'}),
            'ejournal': forms.TextInput(attrs={'class': 'form-control'}),
            'responsible': forms.TextInput(attrs={'class': 'form-control'}),
            'app_version': forms.TextInput(attrs={'class': 'form-control'}),
        }
class TicketCommentForm(forms.ModelForm):
    class Meta:
        model = TicketComment
        fields = ['content']
        widgets = {
            'content': forms.Textarea(attrs={'rows': 4, 'placeholder': 'Add a comment...', 'class': 'form-control'})
        }

class TicketEditForm(forms.ModelForm):
    class Meta:
        model = Ticket
        fields = ['status', 'priority', 'comment_summary', 'problem_category', 'description']
        widgets = {
            'status': forms.Select(attrs={'class': 'form-control'}),
            'priority': forms.Select(attrs={'class': 'form-control'}),
            'comment_summary': forms.Textarea(attrs={'class': 'form-control', 'rows': 2}),
            'problem_category': forms.Select(attrs={'class': 'form-control'}),
            'description': forms.Textarea(attrs={'class': 'form-control', 'rows': 4}),
        }