// vybe-test: csharp/csharp_record_struct_deep/record_struct_bool_ineq
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Flag(bool On); __Check((new Flag(true)==new Flag(false)).ToString(), "False");
