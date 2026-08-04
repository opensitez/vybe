// vybe-test: csharp/csharp_nested_partial_types/nested_interface_is_implemented_by_inner_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

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

class Device {
    public interface IPort { string Open(); }
    public class UsbPort : IPort {
        public string Open() { return "usb-open"; }
    }
}
Device.IPort port = new Device.UsbPort();
__P((port.Open()).ToString());
__Check("usb-open");
