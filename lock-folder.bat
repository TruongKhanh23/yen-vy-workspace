@echo off

icacls "D:\Wife_Workspace" /inheritance:r
icacls "D:\Wife_Workspace" /deny %USERNAME%:(RX,W)

echo LOCKED