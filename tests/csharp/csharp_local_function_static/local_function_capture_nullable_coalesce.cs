// vybe-test: csharp/csharp_local_function_static/local_function_capture_nullable_coalesce
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? maybe=8; int Coalesce(int n){int C(int x)=>x+(maybe??0); return C(n);} __Check((Coalesce(2)).ToString(), "10");
