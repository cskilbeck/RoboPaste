#include <stdio.h>
#include <windows.h>

void Log(WCHAR const *fmt, ...)
{
	va_list v;
	va_start(v, fmt);
	FILE *f;
	if(_wfopen_s(&f, L"D:\\Users\\chs\\Documents\\log.txt", L"a") == 0)
	{
		vfwprintf_s(f, fmt, v);
		fclose(f);
	}
}