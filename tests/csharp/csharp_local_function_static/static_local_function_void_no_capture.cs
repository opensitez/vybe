// vybe-test: csharp/csharp_local_function_static/static_local_function_void_no_capture
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Run(){int acc=0; static void Add(ref int target,int n){target+=n;} Add(ref acc,4); Add(ref acc,1); return acc;} __Check((Run()).ToString(), "5");
