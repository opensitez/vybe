// vybe-test: csharp/csharp_enum_metaprogramming/enum_try_parse_ignore_case_success
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Mode{Alpha,Beta} var ok=System.Enum.TryParse<Mode>("beta",true,out var m); __Check((ok).ToString(), "True"); __Check((m).ToString(), "Beta");
