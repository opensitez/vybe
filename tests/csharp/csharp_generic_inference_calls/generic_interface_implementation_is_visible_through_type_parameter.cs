// vybe-test: csharp/csharp_generic_inference_calls/generic_interface_implementation_is_visible_through_type_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

using static __Harness;

int Load<T>(T device) where T : IReader { return device.Read(); }
__P((Load(new Sensor())).ToString());
__Check("17");

interface IReader {
    int Read();
}

class Sensor : IReader {
    public int Read() { return 17; }
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
