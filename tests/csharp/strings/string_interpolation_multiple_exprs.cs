// vybe-test: csharp/strings/string_interpolation_multiple_exprs
// origin: languages/csharp/tests/csharp/test_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var a = "Alice";
        var age = 30;
        __Check(($"{a} is {age} years old").ToString(), "Alice is 30 years old");
