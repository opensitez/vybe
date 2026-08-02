// vybe-test: csharp/csharp_namespace_aliases/alias_can_reference_nested_class_type
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using InnerType = Demo.Outer.Inner; namespace Demo { public class Outer { public class Inner { public string Name => "inner"; } } } __Check((new InnerType().Name).ToString(), "inner");
