// vybe-test: csharp/csharp_generics_constraints/generic_method_with_interface_constraint_calls_member
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

interface ILabel { string Label(); } class Item : ILabel { public string Label() { return "ok"; } } string Read<T>(T value) where T : ILabel { return value.Label(); } __P((Read(new Item())).ToString());
__Check("ok");
