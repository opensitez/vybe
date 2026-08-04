// vybe-test: csharp/csharp_patterns/method_overloading
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

class Printer {
    public string Print(int x) { return "int:" + x; }
    public string Print(string x) { return "str:" + x; }
    public string Print(int x, int y) { return "pair:" + x + "," + y; }
}
var p = new Printer();
__P((p.Print(42)).ToString());
__P((p.Print("hi")).ToString());
__P((p.Print(1, 2)).ToString());
__Check("int:42\nstr:hi\npair:1,2");
