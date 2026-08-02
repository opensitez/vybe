// vybe-test: csharp/csharp_enum_metaprogramming/enum_parse_whitespace_trim_not_required_exact
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Exact{Ab} var ok=System.Enum.TryParse<Exact>("Ab",out var v); __Check((ok).ToString(), "True");
