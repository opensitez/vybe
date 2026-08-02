// vybe-test: csharp/basics/var_declaration
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 10;
        var y = 20;
        __Check((x + y).ToString(), "30");
