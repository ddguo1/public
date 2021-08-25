@echo.
@echo Allowing remote to remote symlink evaluation
fsutil behavior set SymlinkEvaluation R2R:1
@if errorlevel 1 goto ERROR
@goto END

:ERROR
@echo. 
@echo You need to run this script as Administrator.
@echo Right click on the icon and select "Run as administrator".
@echo. 
@pause

:END
