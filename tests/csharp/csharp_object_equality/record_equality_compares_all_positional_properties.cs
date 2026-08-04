// vybe-test: csharp/csharp_object_equality/record_equality_compares_all_positional_properties
// origin: languages/csharp/tests/csharp/test_csharp_object_equality.rs

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

record Point(int X, int Y);
var a = new Point(1, 2);
var b = new Point(1, 2);
var c = new Point(1, 3);
__P((a.Equals(b)).ToString());
__P((a.Equals(c)).ToString());
__Check("True\nFalse");
