add-type -typeDefinition @"
    using System;
    using System.Runtime.InteropServices;
    using System.Collections.Generic;
    using System.Text;
    using System.Diagnostics;

    namespace windowAPI{
        [Guid("00020893-0000-0000-C000-000000000046")]
        [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
        public interface ExcelWindow
        {
        }

        public class windowApi{
            [DllImport("user32.dll")]
            public static extern IntPtr FindWindow(String sClassName, String sAppName);

            [DllImport("Oleacc.dll")]
            public static extern int AccessibleObjectFromWindow(int hwnd, uint dwObjectID, byte[] riid, out ExcelWindow ptr);

            public static int find(String windowName){
                IntPtr temp;
                temp =  FindWindow(null, windowName);
                return (int)temp;
            }

            public static int accessObject(int hwnd){
                const uint OBJID_NATIVEOM = 0xFFFFFFF0;
                Guid IID_IDispatch = new Guid("{00020400-0000-0000-C000-000000000046}");
                ExcelWindow ptr;

                return AccessibleObjectFromWindow(hwnd, OBJID_NATIVEOM, IID_IDispatch.ToByteArray(), out ptr);
            }
        }
    }
"@

$hwnd = [windowAPI.windowApi]::find("Excel Name.xls")
$excel = [windowAPI.windowApi]::accessObject($hwnd)
