// vybe-test: csharp/csharp_namespace_aliases/qualified_type_name_instantiates_class_from_namespace
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Demo { public class Box { public string Name => "demo"; } } var box = new Demo.Box(); __Check((box.Name).ToString(), "demo");
