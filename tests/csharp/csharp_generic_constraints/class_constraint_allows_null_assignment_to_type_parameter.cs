// vybe-test: csharp/csharp_generic_constraints/class_constraint_allows_null_assignment_to_type_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T AsNull<T>() where T : class => null;
__Check((AsNull<string>() == null).ToString(), "True");
