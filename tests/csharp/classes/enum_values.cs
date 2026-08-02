// vybe-test: csharp/classes/enum_values
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Color { Red, Green, Blue }
        __Check((Color.Red).ToString(), "0");
        __Check((Color.Green).ToString(), "1");
        __Check((Color.Blue).ToString(), "2");
