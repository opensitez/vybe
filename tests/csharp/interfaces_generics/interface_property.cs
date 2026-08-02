// vybe-test: csharp/interfaces_generics/interface_property
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface INamed {
    string Name { get; }
}
class Person : INamed {
    public string Name { get; set; }
}
INamed p = new Person { Name = "Alice" };
__Check((p.Name).ToString(), "Alice");
