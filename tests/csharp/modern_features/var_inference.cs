// vybe-test: csharp/modern_features/var_inference
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

var x = 42;
var s = "hello";
var list = new List<int> { 1, 2, 3 };
__P((x.GetType().Name).ToString());
__P((s.GetType().Name).ToString());
__P((list.Count).ToString());
__Check("Int32\nString\n3");
