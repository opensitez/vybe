// vybe-test: csharp/csharp_structs_value_semantics/nullable_struct_value_exposes_underlying_member
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.DateTime? value = new System.DateTime(2024, 1, 1); __Check((value.Value.Year).ToString(), "2024");
