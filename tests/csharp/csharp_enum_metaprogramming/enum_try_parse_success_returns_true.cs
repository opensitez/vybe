// vybe-test: csharp/csharp_enum_metaprogramming/enum_try_parse_success_returns_true
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Day{Mon,Tue,Wed} var ok=System.Enum.TryParse<Day>("Tue",out var d); __Check((ok).ToString(), "True"); __Check((d).ToString(), "Tue");
