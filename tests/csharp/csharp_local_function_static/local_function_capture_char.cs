// vybe-test: csharp/csharp_local_function_static/local_function_capture_char
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

char ch='A'; string Show(int n){string S(int x)=>ch+""+x; return S(n);} __Check((Show(1)).ToString(), "A1");
