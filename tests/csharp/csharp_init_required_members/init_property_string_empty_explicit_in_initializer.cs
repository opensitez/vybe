// vybe-test: csharp/csharp_init_required_members/init_property_string_empty_explicit_in_initializer
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Label { public string Text { get; init; } = "default"; }
var l = new Label { Text = "" };
__Check((l.Text.Length).ToString(), "0");
