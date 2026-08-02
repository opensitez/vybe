// vybe-test: csharp/csharp_default_interface_methods_deep/default_interface_property_getter
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IProp{int Max=>100;} class Reader:IProp{} __Check((new Reader().Max).ToString(), "100");
