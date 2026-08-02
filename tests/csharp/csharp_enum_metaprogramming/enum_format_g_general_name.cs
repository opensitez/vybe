// vybe-test: csharp/csharp_enum_metaprogramming/enum_format_g_general_name
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Num{X=7} __Check((System.Enum.Format(typeof(Num),Num.X,"G")).ToString(), "X");
