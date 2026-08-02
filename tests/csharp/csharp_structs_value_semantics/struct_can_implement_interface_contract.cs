// vybe-test: csharp/csharp_structs_value_semantics/struct_can_implement_interface_contract
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IText { string Read(); } struct Token : IText { public string Read() { return "ok"; } } IText token = new Token(); __Check((token.Read()).ToString(), "ok");
