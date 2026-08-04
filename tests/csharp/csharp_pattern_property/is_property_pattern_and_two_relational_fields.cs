// vybe-test: csharp/csharp_pattern_property/is_property_pattern_and_two_relational_fields
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

class Range { public int Lo; public int Hi; } object o=new Range{Lo=5,Hi=15}; __P((o is Range{Lo:>0,Hi:<20}).ToString());
__Check("True");
