// vybe-test: csharp/csharp_generics_constraints/generic_method_with_struct_constraint_can_add_nullable_check
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Describe<T>(T? value) where T : struct { return value.HasValue ? value.Value.ToString() : "none"; } __Check((Describe<int>(7)).ToString(), "7");
