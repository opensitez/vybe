// vybe-test: csharp/csharp_record_struct_deep/record_struct_string_case_ineq
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Label(string Text); __Check((new Label("A")==new Label("a")).ToString(), "False");
