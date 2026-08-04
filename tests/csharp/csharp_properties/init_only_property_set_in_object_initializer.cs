// vybe-test: csharp/csharp_properties/init_only_property_set_in_object_initializer
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

class Point { public int X { get; init; } public int Y { get; init; } }
var p = new Point { X=1, Y=2 };
__P((p.X).ToString()); __P((p.Y).ToString());
__Check("1\n2");
