// vybe-test: csharp/csharp_generics_constraints/generic_method_with_new_constraint_constructs_instance
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { public int Value = 9; } T Create<T>() where T : new() { return new T(); } __Check((Create<Box>().Value).ToString(), "9");
