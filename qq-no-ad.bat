@echo off
echo ���� http://www.win7china.com/html/4100.html �������� QQ 2009/2010 ��Ч
echo ���ȹص� QQ �����У�ϵͳ����Ҫ ntfs ��ʽ
for /f "tokens=1,2 delims=:" %%a in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Tencent\QQ2009" /v Install') do ( 
  set "a=%%a"
  set "b=%%b"
)
set "c=%a:~-1%:%b%"
echo --------------------------------------------------------------------------------
echo ��ע����ȡ QQ ��װ·����%c%

echo --------------------------------------------------------------------------------
echo ��ֹ��¼ʱ����������ҳ(������ʧЧ�������в��� http://hi.baidu.com/ddr5707/blog/item/bd89d01b0a1234f0af513311.html)��
set "ad_dll=%c%\Plugin\Com.Tencent.Advertisement\bin\Advertisement.dll"
echo ���� %ad_dll% ֻ������
attrib +r "%ad_dll%"

echo --------------------------------------------------------------------------------
echo ȥ�����촰�����Ͻ�ͼƬ��棺
set "qq_ad_dir=%appdata%\Tencent\QQ\Misc\com.tencent.advertisement"
echo ɾ�� %qq_ad_dir% ������ļ�
del /q %qq_ad_dir%
echo ��ֹ��ǰ�û��� %qq_ad_dir% ��д�ļ�
cacls "%qq_ad_dir%" /e /d "%userdomain%\%username%"
cacls "%qq_ad_dir%" /e /p "%userdomain%\%username%:r"

echo --------------------------------------------------------------------------------
echo ȥ�����촰�����½����ֹ�棺
for /d %%x in ("%appdata%\Tencent\Users\*") do (
  echo ɾ���ļ� %%x\QQ\Misc.db��������ͬ���ļ���
  del /q "%%x\QQ\Misc.db"
  md "%%x\QQ\Misc.db"
)
echo --------------------------------------------------------------------------------
echo ���
pause
