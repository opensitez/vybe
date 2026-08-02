// vybe-test: csharp/csharp_local_function_static/static_local_function_identity
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Id(int n){static int Self(int x)=>x; return Self(n);} __Check((Id(100)).ToString(), "100");
