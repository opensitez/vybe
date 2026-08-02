// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_used_in_string_interpolation
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Tag(string name) { public string Label() => $"tag:{name}"; }
__Check((new Tag("core").Label()).ToString(), "tag:core");
