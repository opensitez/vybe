// vybe-test: csharp/csharp_local_function_static/local_function_used_before_declaration
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((Double(6)).ToString(), "12"); int Double(int x)=>x*2;
