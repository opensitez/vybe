// vybe-test: csharp/csharp_local_function_static/static_local_function_called_from_sibling_local
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Pipeline(int n){static int Double(int x)=>x*2; int Wrap(int v)=>Double(v)+1; return Wrap(n);} __Check((Pipeline(5)).ToString(), "11");
