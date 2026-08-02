// vybe-test: csharp/csharp_local_function_static/local_function_nested_static_inner
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Calc(int n){static int Inner(int x)=>x+5; int Outer(int v)=>Inner(v)*2; return Outer(n);} __Check((Calc(3)).ToString(), "16");
