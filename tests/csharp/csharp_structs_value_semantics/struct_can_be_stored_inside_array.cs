// vybe-test: csharp/csharp_structs_value_semantics/struct_can_be_stored_inside_array
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

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

struct Point { public int X; } var points = new[] { new Point { X = 3 }, new Point { X = 4 } }; foreach (var point in points) __P((point.X).ToString());
__Check("3\n4");
