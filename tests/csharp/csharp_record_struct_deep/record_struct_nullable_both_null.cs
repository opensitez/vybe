// vybe-test: csharp/csharp_record_struct_deep/record_struct_nullable_both_null
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Maybe(int? N); __Check((new Maybe(null)==new Maybe(null)).ToString(), "True");
