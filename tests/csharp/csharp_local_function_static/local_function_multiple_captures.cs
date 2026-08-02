// vybe-test: csharp/csharp_local_function_static/local_function_multiple_captures
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int a=2; int b=3; int Mix(int n){int M(int x)=>a*b+x; return M(n);} __Check((Mix(4)).ToString(), "10");
