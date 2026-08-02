// vybe-test: csharp/csharp_enum_metaprogramming/enum_try_parse_with_ignore_case_false
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Case{Ab} var ok=System.Enum.TryParse<Case>("Ab",false,out var v); __Check((ok).ToString(), "True"); __Check((v).ToString(), "Ab");
