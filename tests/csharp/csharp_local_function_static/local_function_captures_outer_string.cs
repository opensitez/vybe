// vybe-test: csharp/csharp_local_function_static/local_function_captures_outer_string
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string prefix="p:"; string Tag(int n){string T(int x)=>prefix+x; return T(n);} __Check((Tag(7)).ToString(), "p:7");
