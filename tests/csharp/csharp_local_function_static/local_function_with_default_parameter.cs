// vybe-test: csharp/csharp_local_function_static/local_function_with_default_parameter
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Inc(int n){int Step(int x,int by=1)=>x+by; return Step(n,3);} __Check((Inc(10)).ToString(), "13");
