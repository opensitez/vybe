// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_with_enum_parameter
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Mode{On,Off} interface IMode{string Label(Mode m)=>m.ToString();} class M:IMode{} __Check((new M().Label(Mode.On)).ToString(), "On");
