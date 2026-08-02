// vybe-test: csharp/csharp_record_struct_deep/record_struct_long_equal
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Wide(long V); __Check((new Wide(10000000000L)==new Wide(10000000000L)).ToString(), "True");
