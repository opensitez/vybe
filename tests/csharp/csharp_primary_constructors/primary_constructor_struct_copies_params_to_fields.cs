// vybe-test: csharp/csharp_primary_constructors/primary_constructor_struct_copies_params_to_fields
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

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

struct Point(int x, int y) {
    public int X = x;
    public int Y = y;
}
var p = new Point(3, 4);
__P((p.X).ToString()); __P((p.Y).ToString());
__Check("3\n4");
