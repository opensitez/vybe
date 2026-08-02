// vybe-test: csharp/csharp_local_function_static/local_function_captures_mutable_outer
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int scale=2; int Mul(int n){int M(int x)=>x*scale; scale=3; return M(n);} __Check((Mul(4)).ToString(), "12");
