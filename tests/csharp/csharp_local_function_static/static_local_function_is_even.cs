// vybe-test: csharp/csharp_local_function_static/static_local_function_is_even
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool Even(int n){static bool Check(int x)=>x%2==0; return Check(n);} __Check((Even(6)).ToString(), "True");
