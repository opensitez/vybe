// vybe-test: csharp/csharp_const_and_readonly_fields/readonly_struct_field_must_be_set_in_constructor
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

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

struct Cell {
    public readonly int Value;
    public Cell(int value) { Value = value; }
}
__P((new Cell(8).Value).ToString());
__Check("8");
