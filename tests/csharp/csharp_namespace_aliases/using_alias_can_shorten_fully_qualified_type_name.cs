// vybe-test: csharp/csharp_namespace_aliases/using_alias_can_shorten_fully_qualified_type_name
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Thing = Demo.Tools.Box; namespace Demo.Tools { public class Box { public int Value = 7; } } __Check((new Thing().Value).ToString(), "7");
