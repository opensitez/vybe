// vybe-test: csharp/csharp_properties/expression_bodied_property_getter
// origin: languages/csharp/tests/csharp/test_csharp_properties.rs

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

class Rect { public int W,H; public int Area => W * H; }
__P((new Rect{W=3,H=4}.Area).ToString());
__Check("12");
