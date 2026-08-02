// vybe-test: csharp/csharp_local_function_static/static_local_function_recursive_static
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int CountDown(int n){static int Step(int k)=>k<=0?0:1+Step(k-1); return Step(n);} __Check((CountDown(4)).ToString(), "4");
