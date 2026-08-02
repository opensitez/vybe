// vybe-test: csharp/csharp_local_function_static/local_function_if_branch_picks_path
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Sign(int n){string Pos(int x)=>"+"; string Neg(int x)=>"-"; if(n>=0){return Pos(n);} return Neg(n);} __Check((Sign(-1)).ToString(), "-");
