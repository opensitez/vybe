// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_interface_implemented_by_nested_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Device{public interface IPort{string Open();} public class Usb:IPort{public string Open()=>"usb";}} __Check((new Device.Usb().Open()).ToString(), "usb");
