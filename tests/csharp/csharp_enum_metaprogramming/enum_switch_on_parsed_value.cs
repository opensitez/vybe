// vybe-test: csharp/csharp_enum_metaprogramming/enum_switch_on_parsed_value
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Mode{On,Off} var m=(Mode)System.Enum.Parse(typeof(Mode),"On"); string s=m==Mode.On?"yes":"no"; __Check((s).ToString(), "yes");
