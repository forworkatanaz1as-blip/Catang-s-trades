[app]
title = MAGAZIN — Генератор торговцев One Piece
package.name = magazinonepiece
package.domain = org.catang.atanazi
source.dir = .
source.include_exts = py,kv,xlsx,png,jpg,json
source.include_patterns = assets/*
version = 1.0.0
requirements = python3,kivy,openpyxl,pillow
orientation = portrait
fullscreen = 0
android.permissions = READ_EXTERNAL_STORAGE,WRITE_EXTERNAL_STORAGE
android.api = 35
android.minapi = 23
android.archs = arm64-v8a, armeabi-v7a

[buildozer]
log_level = 2
warn_on_root = 1
