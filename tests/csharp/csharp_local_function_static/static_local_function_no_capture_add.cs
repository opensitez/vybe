// vybe-test: csharp/csharp_local_function_static/static_local_function_no_capture_add
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Sum(int a,int b){static int Add(int x,int y)=>x+y; return Add(a,b);} __Check((Sum(3,4)).ToString(), "7");
