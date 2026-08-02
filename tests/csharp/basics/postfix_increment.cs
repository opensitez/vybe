// vybe-test: csharp/basics/postfix_increment
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 5;
        x++;
        x++;
        __Check((x).ToString(), "7");
