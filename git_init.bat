@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist .git (
  echo Git is already initialized.
) else (
  git init
  echo Git repository initialized.
)

echo.
echo Optional: add remote and first commit:
echo   git remote add origin YOUR_REPO_URL
echo   git add .
echo   git commit -m "Initial commit: 오이코스 논문 템플릿 자동 채우기"
echo   git branch -M main
echo   git push -u origin main
echo.
pause
