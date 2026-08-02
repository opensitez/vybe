// vybe-test: csharp/csharp_local_function_static/static_local_function_no_capture_multiply
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Product(int a,int b){static int Mul(int x,int y)=>x*y; return Mul(a,b);} __Check((Product(6,7)).ToString(), "42");
