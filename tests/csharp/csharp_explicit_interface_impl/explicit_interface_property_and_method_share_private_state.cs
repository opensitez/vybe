// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_property_and_method_share_private_state
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

using static __Harness;

IStatus status = new Job();
__P((status.Name).ToString());
__P((status.Read()).ToString());
__Check("queued\nqueued!");

interface IStatus {
    string Name { get; }
    string Read();
}

class Job : IStatus {
    string name = "queued";
    string IStatus.Name { get { return name; } }
    string IStatus.Read() { return name + "!"; }
}

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
