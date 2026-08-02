// vybe-test: csharp/csharp_exception_types/object_disposed_exception_message_is_readable
// origin: languages/csharp/tests/csharp/test_csharp_exception_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try { throw new System.ObjectDisposedException("MyObject"); }
catch(System.ObjectDisposedException e) { __Check((e.ObjectName).ToString(), "MyObject"); }
