// vybe-test: csharp/csharp_local_function_static/static_local_function_min_of_two
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Min(int a,int b){static int Pick(int x,int y)=>x<y?x:y; return Pick(a,b);} __Check((Min(3,9)).ToString(), "3");
