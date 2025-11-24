__version__ = "0.0.1"

# Apply patches on module import
try:
    from export_letterhead.patches import apply_patches
    apply_patches()
except Exception:
    # Silently fail if patches can't be applied (e.g., during installation)
    pass
