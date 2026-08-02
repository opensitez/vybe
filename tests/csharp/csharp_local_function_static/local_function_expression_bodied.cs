// vybe-test: csharp/csharp_local_function_static/local_function_expression_bodied
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Triple(int n){int T(int x)=>x*3; return T(n);} __Check((Triple(5)).ToString(), "15");
