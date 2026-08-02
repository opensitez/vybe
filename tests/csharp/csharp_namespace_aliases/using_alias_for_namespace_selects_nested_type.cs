// vybe-test: csharp/csharp_namespace_aliases/using_alias_for_namespace_selects_nested_type
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Core = Demo.Core; namespace Demo.Core { public class Item { public string Name => "core"; } } __Check((new Core.Item().Name).ToString(), "core");
