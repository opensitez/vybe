// vybe-test: csharp/modern_features/record_equality
// origin: languages/csharp/tests/csharp/test_modern_features.rs

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

record Color(int R, int G, int B);
var c1 = new Color(255, 0, 0);
var c2 = new Color(255, 0, 0);
var c3 = new Color(0, 255, 0);
__P((c1 == c2).ToString());
__P((c1 == c3).ToString());
__Check("True\nFalse");
