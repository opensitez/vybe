// vybe-test: csharp/csharp_enum_metaprogramming/enum_try_parse_empty_string_fails
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Empty{A} var ok=System.Enum.TryParse<Empty>("",out var v); __Check((ok).ToString(), "False");
