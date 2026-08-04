// vybe-test: csharp/csharp_pattern_property/struct_property_pattern_on_value_type
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

struct Vec2 { public int X; public int Y; } object o=new Vec2{X=2,Y=3}; __P((o is Vec2{X:2,Y:3}).ToString());
__Check("True");
