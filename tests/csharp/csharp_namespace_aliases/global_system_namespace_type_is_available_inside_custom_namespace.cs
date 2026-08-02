// vybe-test: csharp/csharp_namespace_aliases/global_system_namespace_type_is_available_inside_custom_namespace
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Demo { public class Worker { public string Read() { return global::System.String.Join(",", new[] { "a", "b" }); } } } __Check((new Demo.Worker().Read()).ToString(), "a,b");
