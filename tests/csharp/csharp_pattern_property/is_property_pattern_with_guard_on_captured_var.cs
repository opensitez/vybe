// vybe-test: csharp/csharp_pattern_property/is_property_pattern_with_guard_on_captured_var
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

class Pair { public int A; public int B; } object o=new Pair{A=4,B=9}; if(o is Pair{A:var a,B:var b} && a<b) __P((b-a).ToString());
__Check("5");
