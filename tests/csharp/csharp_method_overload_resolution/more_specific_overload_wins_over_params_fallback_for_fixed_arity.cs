// vybe-test: csharp/csharp_method_overload_resolution/more_specific_overload_wins_over_params_fallback_for_fixed_arity
// origin: languages/csharp/tests/csharp/test_csharp_method_overload_resolution.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Describe(int value) { return "int:" + value; }
string Describe(params int[] values) { return "many:" + values.Length; }
__Check((Describe(7)).ToString(), "int:7");
__Check((Describe(1, 2)).ToString(), "many:2");
