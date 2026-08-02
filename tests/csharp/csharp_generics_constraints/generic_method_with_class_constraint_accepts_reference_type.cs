// vybe-test: csharp/csharp_generics_constraints/generic_method_with_class_constraint_accepts_reference_type
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Echo<T>(T value) where T : class { return value.ToString(); } __Check((Echo("text")).ToString(), "text");
