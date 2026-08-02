// vybe-test: csharp/csharp_classes/enum_basic
// origin: languages/csharp/tests/csharp/test_csharp_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

enum Color { Red, Green, Blue }
Color c = Color.Green;
__Check(((int)c).ToString(), "1");
