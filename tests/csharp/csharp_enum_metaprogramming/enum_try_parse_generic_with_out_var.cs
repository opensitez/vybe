// vybe-test: csharp/csharp_enum_metaprogramming/enum_try_parse_generic_with_out_var
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Status{Open,Closed} System.Enum.TryParse<Status>("Closed",out var s); __Check((s).ToString(), "Closed");
