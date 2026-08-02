// vybe-test: csharp/csharp_nested_partial_types/partial_class_combines_property_and_constructor_logic
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

partial class Build {
    public string Name { get; set; }
}
partial class Build {
    public Build(string name) { Name = name; }
}
__Check((new Build("nightly").Name).ToString(), "nightly");
