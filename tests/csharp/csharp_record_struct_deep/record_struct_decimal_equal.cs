// vybe-test: csharp/csharp_record_struct_deep/record_struct_decimal_equal
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Money(decimal A); __Check((new Money(9.99m)==new Money(9.99m)).ToString(), "True");
