// vybe-test: csharp/csharp_local_functions/local_function_used_before_its_declaration
// origin: languages/csharp/tests/csharp/test_csharp_local_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((Double(5)).ToString(), "10");
int Double(int x)=>x*2;
