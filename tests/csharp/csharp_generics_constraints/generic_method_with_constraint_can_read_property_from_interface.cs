// vybe-test: csharp/csharp_generics_constraints/generic_method_with_constraint_can_read_property_from_interface
// origin: languages/csharp/tests/csharp/test_csharp_generics_constraints.rs

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

interface INamed { string Name { get; } } class User : INamed { public string Name => "Grace"; } string Read<T>(T item) where T : INamed { return item.Name; } __P((Read(new User())).ToString());
__Check("Grace");
