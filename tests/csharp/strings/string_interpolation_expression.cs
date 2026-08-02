// vybe-test: csharp/strings/string_interpolation_expression
// origin: languages/csharp/tests/csharp/test_strings.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 3;
        var y = 4;
        __Check(($"sum is {x + y}").ToString(), "sum is 7");
