// vybe-test: csharp/csharp_local_function_static/local_function_bool_predicate
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool AllPositive(int a,int b){bool Check(int x,int y)=>x>0&&y>0; return Check(a,b);} __Check((AllPositive(1,2)).ToString(), "True");
