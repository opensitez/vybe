// vybe-test: csharp/csharp_generic_inference_calls/generic_interface_implementation_is_visible_through_type_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
__Check((Load(new Sensor())).ToString(), "17");
