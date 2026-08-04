// vybe-test: csharp/csharp_namespace_aliases/namespace_scoped_interface_is_implemented_by_qualified_type
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

namespace Demo { public interface IRun { string Run(); } public class Worker : IRun { public string Run() { return "done"; } } } Demo.IRun worker = new Demo.Worker(); __P((worker.Run()).ToString());
__Check("done");
