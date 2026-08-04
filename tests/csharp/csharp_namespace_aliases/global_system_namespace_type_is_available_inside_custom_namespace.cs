// vybe-test: csharp/csharp_namespace_aliases/global_system_namespace_type_is_available_inside_custom_namespace
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

namespace Demo { public class Worker { public string Read() { return global::System.String.Join(",", new[] { "a", "b" }); } } } __P((new Demo.Worker().Read()).ToString());
__Check("a,b");
