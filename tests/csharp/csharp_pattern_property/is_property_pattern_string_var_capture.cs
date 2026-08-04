// vybe-test: csharp/csharp_pattern_property/is_property_pattern_string_var_capture
// origin: languages/csharp/tests/csharp/test_csharp_pattern_property.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Label { public string Text; } object o=new Label{Text="go"}; if(o is Label{Text:var t}) __P((t.ToUpper()).ToString());
__Check("GO");
