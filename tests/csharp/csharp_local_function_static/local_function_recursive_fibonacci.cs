// vybe-test: csharp/csharp_local_function_static/local_function_recursive_fibonacci
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Fib(int n){int F(int k)=>k<=1?k:F(k-1)+F(k-2); return F(n);} __Check((Fib(8)).ToString(), "21");
