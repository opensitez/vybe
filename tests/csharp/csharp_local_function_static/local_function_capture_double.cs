// vybe-test: csharp/csharp_local_function_static/local_function_capture_double
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double rate=1.5; int Scale(int n){int S(int x)=>(int)(x*rate); return S(n);} __Check((Scale(4)).ToString(), "6");
