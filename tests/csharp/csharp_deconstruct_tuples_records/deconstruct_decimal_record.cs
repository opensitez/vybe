// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_decimal_record
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Money(decimal A); var (a)=new Money(9.99m); __Check((a).ToString(), "9.99");
