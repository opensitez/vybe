// vybe-test: csharp/csharp_local_function_static/local_function_basic_call_returns_square
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Square(int n){int Sq(int x)=>x*x; return Sq(n);} __Check((Square(4)).ToString(), "16");
