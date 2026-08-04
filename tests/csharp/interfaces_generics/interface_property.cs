// vybe-test: csharp/interfaces_generics/interface_property
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
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
__P((p.Name).ToString());
__Check("Alice");
