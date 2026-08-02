// vybe-test: csharp/csharp_local_function_static/local_function_generic_style_with_object
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Format(int n){string F(int x)=>"n="+x; return F(n);} __Check((Format(42)).ToString(), "n=42");
