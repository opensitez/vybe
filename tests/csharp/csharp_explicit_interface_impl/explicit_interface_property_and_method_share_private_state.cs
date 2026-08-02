// vybe-test: csharp/csharp_explicit_interface_impl/explicit_interface_property_and_method_share_private_state
// origin: languages/csharp/tests/csharp/test_csharp_explicit_interface_impl.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((status.Name).ToString(), "queued");
__Check((status.Read()).ToString(), "queued!");
