// vybe-test: csharp/more_classes/record_tostring
// origin: languages/csharp/tests/csharp/test_more_classes.rs

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

record Point(int X, int Y) {
            public string Display() {
                return "Point(" + X + ", " + Y + ")";
            }
        }
        var p = new Point(3, 7);
        __P((p.Display()).ToString());
__Check("Point(3, 7)");
