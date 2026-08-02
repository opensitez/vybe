// vybe-test: csharp/csharp_local_functions/local_function_returns_func_delegate
// origin: languages/csharp/tests/csharp/test_csharp_local_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,int> MakeAdder(int n){
    int Add(int x)=>x+n;
    return Add;
}
var add10=MakeAdder(10);
__Check((add10(5)).ToString(), "15");
