// vybe-test: csharp/csharp_local_function_static/local_function_capture_long
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

long baseVal=10000000000L; int Add(int n){int A(int x)=>x+(int)(baseVal%100); return A(n);} __Check((Add(5)).ToString(), "5");
