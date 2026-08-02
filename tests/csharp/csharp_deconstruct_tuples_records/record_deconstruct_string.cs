// vybe-test: csharp/csharp_deconstruct_tuples_records/record_deconstruct_string
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Tag(string Name); var (name)=new Tag("z"); __Check((name).ToString(), "z");
