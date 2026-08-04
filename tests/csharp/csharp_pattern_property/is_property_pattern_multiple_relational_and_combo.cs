// vybe-test: csharp/csharp_pattern_property/is_property_pattern_multiple_relational_and_combo
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

class Band { public int Lo; public int Hi; } object o=new Band{Lo=10,Hi=20}; __P((o is Band{Lo:>=10 and <=10,Hi:>=20 and <=20}).ToString());
__Check("True");
