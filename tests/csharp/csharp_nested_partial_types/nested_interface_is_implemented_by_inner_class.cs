// vybe-test: csharp/csharp_nested_partial_types/nested_interface_is_implemented_by_inner_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

using static __Harness;

Device.IPort port = new Device.UsbPort();
__P((port.Open()).ToString());
__Check("usb-open");

class Device {
    public interface IPort { string Open(); }
    public class UsbPort : IPort {
        public string Open() { return "usb-open"; }
    }
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
