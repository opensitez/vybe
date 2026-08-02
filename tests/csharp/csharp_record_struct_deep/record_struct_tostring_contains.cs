// vybe-test: csharp/csharp_record_struct_deep/record_struct_tostring_contains
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Tag(string Name); __Check((new Tag("beta").ToString().Contains("beta")).ToString(), "True");
