[app]
title = ORANLAB Android
package.name = oranlab
package.domain = com.oranlab
source.dir = .
source.include_exts = py,png,jpg,kv,atlas,txt
version = 0.1.0
requirements = python3,kivy,plyer,openpyxl,pyxlsb
orientation = portrait
fullscreen = 0
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.api = 33
android.minapi = 23
android.archs = arm64-v8a, armeabi-v7a
android.allow_backup = False

[buildozer]
log_level = 2
warn_on_root = 1
