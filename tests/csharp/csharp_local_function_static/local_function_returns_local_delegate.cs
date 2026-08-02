// vybe-test: csharp/csharp_local_function_static/local_function_returns_local_delegate
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,int> MakeAdder(int n){int Add(int x)=>x+n; return Add;} var add5=MakeAdder(5); __Check((add5(10)).ToString(), "15");
