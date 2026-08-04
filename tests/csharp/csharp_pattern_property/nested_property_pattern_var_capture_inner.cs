// vybe-test: csharp/csharp_pattern_property/nested_property_pattern_var_capture_inner
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

class Inner { public int N; } class Outer { public Inner Child; } object o=new Outer{Child=new Inner{N=9}}; if(o is Outer{Child:{N:var n}}) __P((n).ToString());
__Check("9");
