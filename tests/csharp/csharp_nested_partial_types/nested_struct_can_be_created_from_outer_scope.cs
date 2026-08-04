// vybe-test: csharp/csharp_nested_partial_types/nested_struct_can_be_created_from_outer_scope
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

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

class Geometry {
    public struct Point {
        public int X;
        public int Y;
    }
}
var point = new Geometry.Point { X = 3, Y = 4 };
__P((point.X + point.Y).ToString());
__Check("7");
