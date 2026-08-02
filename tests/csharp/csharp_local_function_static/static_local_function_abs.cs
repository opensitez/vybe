// vybe-test: csharp/csharp_local_function_static/static_local_function_abs
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Abs(int n){static int Pos(int x)=>x<0?-x:x; return Pos(n);} __Check((Abs(-8)).ToString(), "8");
