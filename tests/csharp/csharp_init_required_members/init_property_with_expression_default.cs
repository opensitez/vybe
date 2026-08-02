// vybe-test: csharp/csharp_init_required_members/init_property_with_expression_default
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Scale { public int Factor { get; init; } = 2 * 3; }
var s = new Scale();
__Check((s.Factor).ToString(), "6");
