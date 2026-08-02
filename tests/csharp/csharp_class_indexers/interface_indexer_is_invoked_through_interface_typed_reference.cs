// vybe-test: csharp/csharp_class_indexers/interface_indexer_is_invoked_through_interface_typed_reference
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((row[1]).ToString(), "b");
