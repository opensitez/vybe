// vybe-test: csharp/csharp_namespace_aliases/using_alias_can_target_generic_type
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using TextList = System.Collections.Generic.List<string>; var list = new TextList { "a", "b" }; __Check((list.Count).ToString(), "2");
