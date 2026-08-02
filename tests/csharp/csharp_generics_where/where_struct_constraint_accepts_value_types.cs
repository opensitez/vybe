// vybe-test: csharp/csharp_generics_where/where_struct_constraint_accepts_value_types
// origin: languages/csharp/tests/csharp/test_csharp_generics_where.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Default<T>() where T:struct=>default;
__Check((Default<int>()).ToString(), "0");
__Check((Default<bool>()).ToString(), "False");
