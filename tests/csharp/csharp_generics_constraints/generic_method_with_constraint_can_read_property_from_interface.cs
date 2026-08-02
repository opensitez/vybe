// vybe-test: csharp/csharp_generics_constraints/generic_method_with_constraint_can_read_property_from_interface
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface INamed { string Name { get; } } class User : INamed { public string Name => "Grace"; } string Read<T>(T item) where T : INamed { return item.Name; } __Check((Read(new User())).ToString(), "Grace");
