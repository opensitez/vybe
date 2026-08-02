// vybe-test: csharp/csharp_default_interface_methods/class_override_replaces_default_interface_method_implementation
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IFormat {
    string Format(int n);
    string Label(int n) { return "d:" + Format(n); }
}
class Custom : IFormat {
    public string Format(int n) { return n.ToString(); }
    public string Label(int n) { return "x:" + Format(n); }
}
IFormat fmt = new Custom();
__Check((fmt.Label(3)).ToString(), "x:3");
