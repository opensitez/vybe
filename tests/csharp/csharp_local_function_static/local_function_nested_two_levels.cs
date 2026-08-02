// vybe-test: csharp/csharp_local_function_static/local_function_nested_two_levels
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Outer(int n){int Mid(int x){int Inner(int y)=>y+1; return Inner(x);} return Mid(n);} __Check((Outer(9)).ToString(), "10");
