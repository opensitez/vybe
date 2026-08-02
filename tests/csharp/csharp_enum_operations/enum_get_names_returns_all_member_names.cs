// vybe-test: csharp/csharp_enum_operations/enum_get_names_returns_all_member_names
// origin: languages/csharp/tests/csharp/test_csharp_enum_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Coin{Penny,Nickel,Dime}
__Check((System.Enum.GetNames(typeof(Coin)).Length).ToString(), "3");
