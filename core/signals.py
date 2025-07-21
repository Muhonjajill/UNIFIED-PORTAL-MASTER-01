# core/signals.py
from django.db.models.signals import post_migrate, post_save
from django.dispatch import receiver
from django.contrib.auth.models import User, Group, Permission
from django.contrib.contenttypes.models import ContentType
from core.models import File, Profile
from django.core.exceptions import ObjectDoesNotExist


@receiver(post_migrate)
def setup_groups_and_permissions(sender, **kwargs):
    from django.core.exceptions import ObjectDoesNotExist

    # File model permissions
    file_content_type = ContentType.objects.get_for_model(File)
    user_content_type = ContentType.objects.get_for_model(User)

    # Create groups
    admin_group, _ = Group.objects.get_or_create(name='Admin')
    editor_group, _ = Group.objects.get_or_create(name='Editor')
    viewer_group, _ = Group.objects.get_or_create(name='Viewer')

    try:
        # File permissions
        view_file = Permission.objects.get(codename='view_file')
        change_file = Permission.objects.get(codename='change_file')
        delete_file = Permission.objects.get(codename='delete_file')
        add_file = Permission.objects.get(codename='add_file')  

        can_edit_file_perm, _ = Permission.objects.get_or_create(
            codename='can_edit_file',
            name='Can edit file',
            content_type=file_content_type
        )

        # User model permissions
        view_user = Permission.objects.get(codename='view_user')
        change_user = Permission.objects.get(codename='change_user')
        delete_user = Permission.objects.get(codename='delete_user')

    except ObjectDoesNotExist:
        return

    # Assign permissions to groups
    admin_group.permissions.set([
        view_file, change_file, delete_file, add_file, can_edit_file_perm,
        view_user, change_user, delete_user
    ])

    editor_group.permissions.set([
        view_file, change_file, add_file, 
        view_user, change_user
    ])

    viewer_group.permissions.set([
        view_file,
        view_user
    ])


@receiver(post_save, sender=User)
def manage_user_profile(sender, instance, created, **kwargs):
    if created:
        Profile.objects.create(user=instance)
        if instance.is_superuser:
            admin_group, _ = Group.objects.get_or_create(name='Admin')
            instance.groups.add(admin_group)
    else:
        if hasattr(instance, 'profile'):
            instance.profile.save()
