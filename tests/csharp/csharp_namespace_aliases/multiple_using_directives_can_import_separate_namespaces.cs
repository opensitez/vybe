// vybe-test: csharp/csharp_namespace_aliases/multiple_using_directives_can_import_separate_namespaces
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo.Left; using Demo.Right; namespace Demo.Left { public class A { public string Name => "A"; } } namespace Demo.Right { public class B { public string Name => "B"; } } __Check((new A().Name + new B().Name).ToString(), "AB");
