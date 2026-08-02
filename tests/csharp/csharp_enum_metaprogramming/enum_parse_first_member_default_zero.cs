// vybe-test: csharp/csharp_enum_metaprogramming/enum_parse_first_member_default_zero
// origin: languages/csharp/tests/csharp/test_csharp_enum_metaprogramming.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum State{Idle,Running} var s=(State)System.Enum.Parse(typeof(State),"Idle"); __Check(((int)s).ToString(), "0");
