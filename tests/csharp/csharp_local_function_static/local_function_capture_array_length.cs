// vybe-test: csharp/csharp_local_function_static/local_function_capture_array_length
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={1,2,3}; int LenPlus(int n){int L(int x)=>data.Length+x; return L(n);} __Check((LenPlus(2)).ToString(), "5");
