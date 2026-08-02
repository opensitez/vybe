// vybe-test: csharp/csharp_nullable_value_deep/nullable_int_lifted_division_both_present
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int? a=20; int? b=4; __Check((a/b).ToString(), "5");
