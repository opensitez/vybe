// vybe-test: csharp/csharp_local_function_static/local_function_tail_call_style_sum
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Sum(int n){int Loop(int i,int acc)=>i>n?acc:Loop(i+1,acc+i); return Loop(1,0);} __Check((Sum(4)).ToString(), "10");
