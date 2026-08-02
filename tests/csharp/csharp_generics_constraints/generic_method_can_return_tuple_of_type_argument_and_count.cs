// vybe-test: csharp/csharp_generics_constraints/generic_method_can_return_tuple_of_type_argument_and_count
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

(T, int) Pair<T>(T value) { return (value, 1); } var result = Pair("x"); __Check((result.Item1 + result.Item2).ToString(), "x1");
