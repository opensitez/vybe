// vybe-test: csharp/csharp_generic_constraints/struct_constraint_produces_non_null_default_value
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Zero<T>() where T : struct => default;
__Check((Zero<int>()).ToString(), "0");
