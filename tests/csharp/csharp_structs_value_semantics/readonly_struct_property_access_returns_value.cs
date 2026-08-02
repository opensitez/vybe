// vybe-test: csharp/csharp_structs_value_semantics/readonly_struct_property_access_returns_value
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly struct Size { public int Width { get; } public Size(int width) { Width = width; } } __Check((new Size(6).Width).ToString(), "6");
