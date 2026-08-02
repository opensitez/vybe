// vybe-test: csharp/more_classes/as_cast
// origin: languages/csharp/tests/csharp/test_more_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object x = "hello";
        var s = x as string;
        __Check((s).ToString(), "hello");
