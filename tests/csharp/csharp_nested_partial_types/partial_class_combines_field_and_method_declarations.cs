// vybe-test: csharp/csharp_nested_partial_types/partial_class_combines_field_and_method_declarations
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

partial class Config {
    string env = "prod";
}
partial class Config {
    public string Read() { return env; }
}
__Check((new Config().Read()).ToString(), "prod");
