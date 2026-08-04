// vybe-test: csharp/common_patterns/ref_parameter
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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

class Ops {
    public static void Double(ref int x) { x *= 2; }
}
int val = 5;
Ops.Double(ref val);
__P((val).ToString());
Ops.Double(ref val);
__P((val).ToString());
__Check("10\n20");
