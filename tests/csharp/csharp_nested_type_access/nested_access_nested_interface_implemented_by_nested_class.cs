// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_interface_implemented_by_nested_class
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

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

class Device{public interface IPort{string Open();} public class Usb:IPort{public string Open()=>"usb";}} __P((new Device.Usb().Open()).ToString());
__Check("usb");
