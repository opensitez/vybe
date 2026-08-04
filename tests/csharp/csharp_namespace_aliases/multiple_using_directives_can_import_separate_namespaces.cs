// vybe-test: csharp/csharp_namespace_aliases/multiple_using_directives_can_import_separate_namespaces
// origin: languages/csharp/tests/csharp/test_csharp_namespace_aliases.rs

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

using Demo.Left; using Demo.Right; namespace Demo.Left { public class A { public string Name => "A"; } } namespace Demo.Right { public class B { public string Name => "B"; } } __P((new A().Name + new B().Name).ToString());
__Check("AB");
