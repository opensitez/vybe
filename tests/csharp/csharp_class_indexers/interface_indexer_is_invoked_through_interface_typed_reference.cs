// vybe-test: csharp/csharp_class_indexers/interface_indexer_is_invoked_through_interface_typed_reference
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

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

interface ICell {
    string this[int index] { get; }
}
class Row : ICell {
    string[] cells = { "a", "b" };
    public string this[int index] { get { return cells[index]; } }
}
ICell row = new Row();
__P((row[1]).ToString());
__Check("b");
