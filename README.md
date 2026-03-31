# 华美物流发货单 - HarmonyOS APK

## 功能说明
- 本APP为WebView包装壳，需要配合电脑端Web服务使用
- 启动电脑端PWA服务后，手机APP自动连接

## 使用方法
1. 在电脑上启动PWA服务：`python3 app.py`（在PWA版目录）
2. 手机和电脑在同一WiFi网络
3. 查看电脑IP后，修改`MainActivity.java`中的IP地址
4. 安装APK

## 构建APK
```bash
# 本地构建需要Android SDK
./gradlew assembleDebug
```
