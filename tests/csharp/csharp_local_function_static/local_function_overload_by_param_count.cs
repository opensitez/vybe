// vybe-test: csharp/csharp_local_function_static/local_function_overload_by_param_count
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Compute(int n){int One(int x)=>x+1; int Two(int x,int y)=>x+y; return Two(n,One(n));} __Check((Compute(5)).ToString(), "11");
