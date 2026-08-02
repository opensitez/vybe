// vybe-test: csharp/csharp_generics_constraints/generic_method_can_compare_two_values_with_equality
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

bool Same<T>(T left, T right) { return left.Equals(right); } __Check((Same(3, 3)).ToString(), "True");
