// vybe-test: csharp/csharp_local_function_static/local_function_captures_outer_int
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int offset=10; int Add(int n){int B(int x)=>x+offset; return B(n);} __Check((Add(5)).ToString(), "15");
