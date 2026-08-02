// vybe-test: csharp/csharp_enum_metaprogramming/enum_parse_case_sensitive_by_default
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Case{Ab} var ok=System.Enum.TryParse<Case>("ab",out var v); __Check((ok).ToString(), "False");
