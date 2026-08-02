// vybe-test: csharp/csharp_const_and_readonly_fields/readonly_struct_field_must_be_set_in_constructor
// origin: languages/csharp/tests/csharp/test_csharp_const_and_readonly_fields.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Cell {
    public readonly int Value;
    public Cell(int value) { Value = value; }
}
__Check((new Cell(8).Value).ToString(), "8");
