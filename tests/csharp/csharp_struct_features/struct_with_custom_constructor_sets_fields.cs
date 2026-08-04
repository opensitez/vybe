// vybe-test: csharp/csharp_struct_features/struct_with_custom_constructor_sets_fields
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

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

struct Rect { public int W,H; public Rect(int w, int h) { W=w; H=h; } }
var r = new Rect(3, 4);
__P((r.W * r.H).ToString());
__Check("12");
