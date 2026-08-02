// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_used_in_indexer
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Row(int size) {
    int[] cells = new int[size];
    public int this[int i] { get => cells[i]; set => cells[i] = value; }
}
var r = new Row(3);
r[1] = 9;
__Check((r[1]).ToString(), "9");
