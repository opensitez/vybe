// vybe-test: csharp/csharp_nested_partial_types/nested_interface_is_implemented_by_inner_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((port.Open()).ToString(), "usb-open");
