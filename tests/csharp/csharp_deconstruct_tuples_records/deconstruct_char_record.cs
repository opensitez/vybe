// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_char_record
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Sym(char C); var (c)=new Sym('Q'); __Check((c).ToString(), "Q");
