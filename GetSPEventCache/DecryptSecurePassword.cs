using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace SPEventCache
{
    class DecryptSecurePassword
    {
        internal static string SecurePasswordToString(SecureString secureStringPassword)
        {
            var password = string.Empty;
            var intPtrZero = IntPtr.Zero;

            try
            {
                intPtrZero = Marshal.SecureStringToBSTR(secureStringPassword);
                password = Marshal.PtrToStringBSTR(intPtrZero);
            }
            finally
            {
                if (IntPtr.Zero != intPtrZero)
                {
                    Marshal.ZeroFreeBSTR(intPtrZero);
                }
            }
            return password;
        }
    }
}
