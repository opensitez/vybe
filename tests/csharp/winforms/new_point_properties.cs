// vybe-test: csharp/winforms/new_point_properties
// origin: languages/csharp/tests/csharp/test_winforms.rs

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

var p = new Point(100, 200);
        __P((p.x).ToString());
        __P((p.y).ToString());
__Check("100\n200");
