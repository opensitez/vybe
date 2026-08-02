// vybe-test: csharp/csharp_record_struct_deep/record_struct_single_field_equal
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Count(int N); __Check((new Count(0)==new Count(0)).ToString(), "True");
