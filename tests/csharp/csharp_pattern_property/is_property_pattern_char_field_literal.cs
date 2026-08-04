// vybe-test: csharp/csharp_pattern_property/is_property_pattern_char_field_literal
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

class Glyph { public char Ch; } object o=new Glyph{Ch='Z'}; __P((o is Glyph{Ch:'Z'}).ToString());
__Check("True");
