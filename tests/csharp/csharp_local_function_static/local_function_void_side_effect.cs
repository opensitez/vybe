// vybe-test: csharp/csharp_local_function_static/local_function_void_side_effect
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Run(){int acc=0; void Bump(int n){acc+=n;} Bump(2); Bump(3); return acc;} __Check((Run()).ToString(), "5");
