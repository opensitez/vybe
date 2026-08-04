// vybe-test: csharp/csharp_generic_inference_calls/generic_interface_implementation_is_visible_through_type_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

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

interface IReader {
    int Read();
}
class Sensor : IReader {
    public int Read() { return 17; }
}
int Load<T>(T device) where T : IReader { return device.Read(); }
__P((Load(new Sensor())).ToString());
__Check("17");
