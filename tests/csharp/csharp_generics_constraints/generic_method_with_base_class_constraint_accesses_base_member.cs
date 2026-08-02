// vybe-test: csharp/csharp_generics_constraints/generic_method_with_base_class_constraint_accesses_base_member
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { public string Name = "base"; } class Child : Base { } string Read<T>(T value) where T : Base { return value.Name; } __Check((Read(new Child())).ToString(), "base");
