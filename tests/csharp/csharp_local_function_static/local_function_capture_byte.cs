// vybe-test: csharp/csharp_local_function_static/local_function_capture_byte
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

byte mask=3; int Apply(int n){int A(int x)=>x+(int)mask; return A(n);} __Check((Apply(5)).ToString(), "8");
