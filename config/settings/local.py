from .base import *

DEBUG = True

ALLOWED_HOSTS = ["*"]

DB_NAME = config("DB_NAME", default="")

if DB_NAME:
    DATABASES = {
        "default": {
            "ENGINE": "django.db.backends.postgresql",
            "NAME": DB_NAME,
            "USER": config("DB_USER", default="postgres"),
            "PASSWORD": config("DB_PASSWORD", default=""),
            "HOST": config("DB_HOST", default="localhost"),
            "PORT": config("DB_PORT", default="5432"),
            "OPTIONS": {
                "sslmode": config("DB_SSLMODE", default="disable"),
            } if config("DB_SSLMODE", default="") else {},
        }
    }
else:
    DATABASES = {
        "default": {
            "ENGINE": "django.db.backends.sqlite3",
            "NAME": BASE_DIR / "db.sqlite3",
        }
    }

CORS_ALLOW_ALL_ORIGINS = True

# Email
EMAIL_BACKEND = config("EMAIL_BACKEND", default="django.core.mail.backends.console.EmailBackend")
EMAIL_HOST = config("EMAIL_HOST", default="smtp.gmail.com")
EMAIL_PORT = config("EMAIL_PORT", default=587, cast=int)
EMAIL_USE_TLS = config("EMAIL_USE_TLS", default=True, cast=bool)
EMAIL_USE_SSL = config("EMAIL_USE_SSL", default=False, cast=bool)
EMAIL_TIMEOUT = config("EMAIL_TIMEOUT", default=10, cast=int)
EMAIL_HOST_USER = config("EMAIL_HOST_USER", default="solodevbuilds@gmail.com")
EMAIL_HOST_PASSWORD = config("EMAIL_HOST_PASSWORD", default="mcve rnxb rnrz kxuy")
DEFAULT_FROM_EMAIL = config("DEFAULT_FROM_EMAIL", default="Horizon Industrial Parks <solodevbuilds@gmail.com>")

# Vendor portal
VENDOR_PORTAL_BASE_URL = "http://localhost:8080"
