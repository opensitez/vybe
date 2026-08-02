// vybe-test: csharp/csharp_pattern_property/is_property_pattern_string_var_capture
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Label { public string Text; } object o=new Label{Text="go"}; if(o is Label{Text:var t}) __Check((t.ToUpper()).ToString(), "GO");
