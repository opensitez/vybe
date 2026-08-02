// vybe-test: csharp/csharp_local_functions/static_local_function_cannot_capture_outer_variable
// origin: languages/csharp/tests/csharp/test_csharp_local_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static int Pure(int a,int b){
    static int Add(int x,int y)=>x+y;
    return Add(a,b);
}
__Check((Pure(4,5)).ToString(), "9");
