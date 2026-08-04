// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_property_and_method_share_private_state
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

interface IStatus {
    string Name { get; }
    string Read();
}
class Job : IStatus {
    string name = "queued";
    string IStatus.Name { get { return name; } }
    string IStatus.Read() { return name + "!"; }
}
IStatus status = new Job();
__P((status.Name).ToString());
__P((status.Read()).ToString());
__Check("queued\nqueued!");
