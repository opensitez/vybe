// vybe-test: csharp/csharp_string_interpolation/simple_variable_interpolation_in_string
// origin: languages/csharp/tests/csharp/test_csharp_string_interpolation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string name="World"; __Check(($"Hello {name}!").ToString(), "Hello World!");
