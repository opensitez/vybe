// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_used_in_indexer
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

class Row(int size) {
    int[] cells = new int[size];
    public int this[int i] { get => cells[i]; set => cells[i] = value; }
}
var r = new Row(3);
r[1] = 9;
__P((r[1]).ToString());
__Check("9");
