// vybe-test: csharp/csharp_generics_constraints/generic_method_with_struct_constraint_accepts_value_type
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Wrap<T>(T value) where T : struct { return value.ToString(); } __Check((Wrap(5)).ToString(), "5");
