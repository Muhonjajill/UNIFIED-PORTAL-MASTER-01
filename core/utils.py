# core/utils.py
def is_admin(user):
    return user.is_superuser or user.groups.filter(name='Admin').exists()

def is_editor(user):
    return user.groups.filter(name='Editor').exists()

def is_viewer(user):
    return user.groups.filter(name='Viewer').exists()

def can_user_access_file(user, file):
    if file.access_level == 'admin' and not is_admin(user):
        return False
    elif file.access_level == 'editor' and not (is_admin(user) or is_editor(user)):
        return False
    elif file.access_level == 'viewer':
        return True
    return True
