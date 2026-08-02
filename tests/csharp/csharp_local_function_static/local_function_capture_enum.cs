// vybe-test: csharp/csharp_local_function_static/local_function_capture_enum
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Mode{On,Off} Mode state=Mode.On; int Code(int n){int C(int x)=>state==Mode.On?x:0; return C(n);} __Check((Code(5)).ToString(), "5");
