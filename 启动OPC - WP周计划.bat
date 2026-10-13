@echo off

REM 1. 切换到项目目录
cd /d "C:\AIProjects\3-wp-weekly-scheduling"
REM 2. 激活虚拟环境（多个项目共用这一个venv，不用改）
call "C:\MyPythonVenvs\contract_extract_venv\Scripts\activate"
REM 3. 启动OpenCode，加载这个项目对应的会话，替换成你的账单sessionID
opencode -s "ses_ed96b926fffeT117NlJKx152fX"
REM 4. 暂停看报错
pause