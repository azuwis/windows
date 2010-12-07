@echo off
echo 根据 http://www.win7china.com/html/4100.html 制作，对 QQ 2009/2010 有效
echo 请先关掉 QQ 再运行，系统盘需要 ntfs 格式
for /f "tokens=1,2 delims=:" %%a in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Tencent\QQ2009" /v Install') do ( 
  set "a=%%a"
  set "b=%%b"
)
set "c=%a:~-1%:%b%"
echo --------------------------------------------------------------------------------
echo 从注册表读取 QQ 安装路径：%c%

echo --------------------------------------------------------------------------------
echo 禁止登录时弹出迷你首页(方法已失效，请另行参照 http://hi.baidu.com/ddr5707/blog/item/bd89d01b0a1234f0af513311.html)：
set "ad_dll=%c%\Plugin\Com.Tencent.Advertisement\bin\Advertisement.dll"
echo 设置 %ad_dll% 只读属性
attrib +r "%ad_dll%"

echo --------------------------------------------------------------------------------
echo 去掉聊天窗口右上角图片广告：
set "qq_ad_dir=%appdata%\Tencent\QQ\Misc\com.tencent.advertisement"
echo 删除 %qq_ad_dir% 里面的文件
del /q %qq_ad_dir%
echo 禁止当前用户往 %qq_ad_dir% 里写文件
cacls "%qq_ad_dir%" /e /d "%userdomain%\%username%"
cacls "%qq_ad_dir%" /e /p "%userdomain%\%username%:r"

echo --------------------------------------------------------------------------------
echo 去掉聊天窗口左下角文字广告：
for /d %%x in ("%appdata%\Tencent\Users\*") do (
  echo 删除文件 %%x\QQ\Misc.db，并创建同名文件夹
  del /q "%%x\QQ\Misc.db"
  md "%%x\QQ\Misc.db"
)
echo --------------------------------------------------------------------------------
echo 完成
pause
