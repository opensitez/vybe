// vybe-test: csharp/csharp_collection_initializer_syntax/object_initializer_sets_public_fields_before_use
// origin: languages/csharp/tests/csharp/test_csharp_collection_initializer_syntax.rs

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

class Point { public int X; public int Y; }
var point = new Point { X = 2, Y = 5 };
__P((point.X + point.Y).ToString());
__Check("7");
