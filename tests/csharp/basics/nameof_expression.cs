// vybe-test: csharp/basics/nameof_expression
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var myVar = 42;
        __Check((nameof(myVar)).ToString(), "myVar");
