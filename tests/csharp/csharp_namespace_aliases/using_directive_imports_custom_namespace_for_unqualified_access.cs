// vybe-test: csharp/csharp_namespace_aliases/using_directive_imports_custom_namespace_for_unqualified_access
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo.Tools; namespace Demo.Tools { public class Worker { public string Name => "tool"; } } __Check((new Worker().Name).ToString(), "tool");
