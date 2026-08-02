// vybe-test: csharp/csharp_record_struct_deep/record_struct_negative_equal
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Offset(int Delta); __Check((new Offset(-5)==new Offset(-5)).ToString(), "True");
