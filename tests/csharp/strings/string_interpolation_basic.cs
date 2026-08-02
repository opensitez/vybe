// vybe-test: csharp/strings/string_interpolation_basic
// origin: languages/csharp/tests/csharp/test_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var name = "World";
        __Check(($"Hello {name}!").ToString(), "Hello World!");
