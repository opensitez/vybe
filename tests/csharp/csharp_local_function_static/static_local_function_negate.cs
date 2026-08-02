// vybe-test: csharp/csharp_local_function_static/static_local_function_negate
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Negate(int n){static int Flip(int x)=>-x; return Flip(n);} __Check((Negate(12)).ToString(), "-12");
