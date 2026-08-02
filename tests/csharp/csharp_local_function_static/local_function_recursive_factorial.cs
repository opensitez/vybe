// vybe-test: csharp/csharp_local_function_static/local_function_recursive_factorial
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Fact(int n){int F(int k)=>k<=1?1:k*F(k-1); return F(n);} __Check((Fact(5)).ToString(), "120");
