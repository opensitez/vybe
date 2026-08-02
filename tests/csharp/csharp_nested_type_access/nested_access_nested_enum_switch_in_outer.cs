// vybe-test: csharp/csharp_nested_type_access/nested_access_nested_enum_switch_in_outer
// origin: languages/csharp/tests/csharp/test_csharp_nested_type_access.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Gate{public enum Mode{On,Off} public string Label(Mode m){switch(m){case Mode.On:return "on"; default:return "off";}}} __Check((new Gate().Label(Gate.Mode.On)).ToString(), "on");
