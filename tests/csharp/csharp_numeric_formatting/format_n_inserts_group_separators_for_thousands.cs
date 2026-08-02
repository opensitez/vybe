// vybe-test: csharp/csharp_numeric_formatting/format_n_inserts_group_separators_for_thousands
// origin: languages/csharp/tests/csharp/test_csharp_numeric_formatting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var s = (1234567).ToString("N0",
    System.Globalization.CultureInfo.InvariantCulture);
__Check((s).ToString(), "1,234,567");
