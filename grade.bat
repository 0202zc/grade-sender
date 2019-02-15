@echo off 
if "%1" == "h" goto begin 
mshta vbscript:createobject("wscript.shell").run("""%~nx0"" h",0)(window.close)&&exit 
:begin
python "C:\\Users\\Administrator\\Desktop\\成绩监控并推送(手机输入验证码).py"