// vybe-test: csharp/csharp_type_conversions/as_operator_returns_null_for_non_matching_type
// origin: languages/csharp/tests/csharp/test_csharp_type_conversions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object item = 42; string text = item as string; __Check((text is null).ToString(), "True");
