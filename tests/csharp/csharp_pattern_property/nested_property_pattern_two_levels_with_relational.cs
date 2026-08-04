// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_two_levels_with_relational
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

class Inner { public int N; } class Outer { public Inner I; } object o=new Outer{I=new Inner{N=50}}; __P((o is Outer{I:{N:>40}}).ToString());
__Check("True");
