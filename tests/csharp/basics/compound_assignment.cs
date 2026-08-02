// vybe-test: csharp/basics/compound_assignment
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 10;
        x += 5;
        x -= 3;
        x *= 2;
        __Check((x).ToString(), "24");
