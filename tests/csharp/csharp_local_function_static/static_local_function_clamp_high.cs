// vybe-test: csharp/csharp_local_function_static/static_local_function_clamp_high
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Clamp(int n,int max){static int Cap(int x,int m)=>x>m?m:x; return Cap(n,max);} __Check((Clamp(15,10)).ToString(), "10");
