// vybe-test: csharp/csharp_structs_value_semantics/struct_method_can_compute_from_fields
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

struct Point { public int X; public int Y; public int Sum() { return X + Y; } } var point = new Point { X = 4, Y = 6 }; __P((point.Sum()).ToString());
__Check("10");
