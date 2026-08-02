// vybe-test: csharp/csharp_enum_operations/flags_enum_combined_value_has_expected_integer
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

[System.Flags] enum Perm{None=0,Read=1,Write=2,Execute=4}
__Check(((int)(Perm.Read|Perm.Execute)).ToString(), "5");
