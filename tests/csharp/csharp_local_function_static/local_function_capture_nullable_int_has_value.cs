// vybe-test: csharp/csharp_local_function_static/local_function_capture_nullable_int_has_value
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? opt=7; int Bump(int n){int B(int x)=>x+(opt??0); return B(n);} __Check((Bump(1)).ToString(), "8");
