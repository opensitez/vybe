// vybe-test: csharp/csharp_generics_constraints/generic_method_with_interface_constraint_calls_member
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ILabel { string Label(); } class Item : ILabel { public string Label() { return "ok"; } } string Read<T>(T value) where T : ILabel { return value.Label(); } __Check((Read(new Item())).ToString(), "ok");
