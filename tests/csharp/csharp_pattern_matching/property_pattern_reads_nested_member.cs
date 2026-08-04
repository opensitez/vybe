// vybe-test: csharp/csharp_pattern_matching/property_pattern_reads_nested_member
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

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

class Rect { public int W, H; }
object r = new Rect { W=10, H=5 };
string size = r switch { Rect { W: > 8 } => "wide", _ => "narrow" };
__P((size).ToString());
__Check("wide");
