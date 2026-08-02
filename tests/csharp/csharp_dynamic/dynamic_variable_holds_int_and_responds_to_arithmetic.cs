// vybe-test: csharp/csharp_dynamic/dynamic_variable_holds_int_and_responds_to_arithmetic
// origin: languages/csharp/tests/csharp/test_csharp_dynamic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

dynamic x=5; x+=3;
__Check((x).ToString(), "8");
