// vybe-test: csharp/csharp_raw_string_literals/raw_string_compare_ordinal_ignore_case
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string a="""Hello"""; string b="""hello"""; __Check((string.Equals(a,b,System.StringComparison.OrdinalIgnoreCase)).ToString(), "True");
