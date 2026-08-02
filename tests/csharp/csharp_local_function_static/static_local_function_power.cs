// vybe-test: csharp/csharp_local_function_static/static_local_function_power
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Pow(int b,int e){static int Loop(int base,int exp,int acc)=>exp==0?acc:Loop(base,exp-1,acc*base); return Loop(b,e,1);} __Check((Pow(2,4)).ToString(), "16");
