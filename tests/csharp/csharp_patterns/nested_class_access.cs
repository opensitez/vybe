// vybe-test: csharp/csharp_patterns/nested_class_access
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

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

class Outer {
    public int Value = 10;
    public class Inner {
        public int Value = 20;
    }
}
var o = new Outer();
var i = new Outer.Inner();
__P((o.Value).ToString());
__P((i.Value).ToString());
__Check("10\n20");
