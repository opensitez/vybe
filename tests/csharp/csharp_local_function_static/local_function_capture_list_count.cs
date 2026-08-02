// vybe-test: csharp/csharp_local_function_static/local_function_capture_list_count
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var items=new System.Collections.Generic.List<int>{1,2,3}; int SizePlus(int n){int S(int x)=>items.Count+x; return S(n);} __Check((SizePlus(1)).ToString(), "4");
