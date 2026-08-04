// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_method_is_invoked_after_cast_from_object
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

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

interface IRunner { string Run(); }
class TaskRunner : IRunner {
    string IRunner.Run() { return "done"; }
}
object item = new TaskRunner();
__P((((IRunner)item).Run()).ToString());
__Check("done");
