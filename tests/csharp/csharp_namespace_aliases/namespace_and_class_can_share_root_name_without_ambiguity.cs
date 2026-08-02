// vybe-test: csharp/csharp_namespace_aliases/namespace_and_class_can_share_root_name_without_ambiguity
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Demo.Sub { public class Demo { public int Value = 5; } } __Check((new Demo.Sub.Demo().Value).ToString(), "5");
