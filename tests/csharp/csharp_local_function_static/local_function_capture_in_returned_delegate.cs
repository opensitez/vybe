// vybe-test: csharp/csharp_local_function_static/local_function_capture_in_returned_delegate
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Func<int,int> MakeScaler(int factor){int Scale(int x)=>x*factor; return Scale;} __Check((MakeScaler(4)(6)).ToString(), "24");
