// vybe-test: csharp/csharp_namespace_aliases/distinct_namespaces_can_define_same_type_name
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

namespace Left { public class Item { public string Name => "L"; } } namespace Right { public class Item { public string Name => "R"; } } __Check((new Left.Item().Name).ToString(), "L"); __Check((new Right.Item().Name).ToString(), "R");
