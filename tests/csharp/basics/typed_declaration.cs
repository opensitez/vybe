// vybe-test: csharp/basics/typed_declaration
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 5;
        double y = 3.14;
        __Check((x).ToString(), "5");
        __Check((y).ToString(), "3.14");
