// vybe-test: csharp/csharp_enum_metaprogramming/enum_try_parse_failure_returns_false
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Day{Mon,Tue} var ok=System.Enum.TryParse<Day>("Sun",out var d); __Check((ok).ToString(), "False");
