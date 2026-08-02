// vybe-test: csharp/csharp_namespace_aliases/namespace_scoped_enum_is_resolved_by_qualified_name
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Demo { public enum State { Ready } } __Check((Demo.State.Ready).ToString(), "Ready");
