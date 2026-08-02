// vybe-test: csharp/csharp_namespace_aliases/nested_namespace_type_is_reachable_by_full_name
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Outer.Inner { public class Worker { public string Run() { return "ok"; } } } __Check((new Outer.Inner.Worker().Run()).ToString(), "ok");
