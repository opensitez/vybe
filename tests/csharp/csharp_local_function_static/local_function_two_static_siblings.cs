// vybe-test: csharp/csharp_local_function_static/local_function_two_static_siblings
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Combo(int n){static int A(int x)=>x+1; static int B(int x)=>x*2; return B(A(n));} __Check((Combo(4)).ToString(), "10");
