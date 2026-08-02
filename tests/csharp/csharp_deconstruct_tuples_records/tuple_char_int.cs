// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_char_int
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var (ch,n)=('A',1); __Check((ch).ToString(), "A"); __Check((n).ToString(), "1");
