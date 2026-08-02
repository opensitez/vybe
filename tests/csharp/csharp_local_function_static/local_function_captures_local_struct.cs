// vybe-test: csharp/csharp_local_function_static/local_function_captures_local_struct
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int UseStruct(){var p=new System.ValueTuple<int,int>(2,3); int Sum(int n){int S(int x)=>p.Item1+p.Item2+x; return S(n);} return Sum(1);} __Check((UseStruct()).ToString(), "6");
