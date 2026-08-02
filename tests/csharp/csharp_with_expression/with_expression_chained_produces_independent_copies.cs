// vybe-test: csharp/csharp_with_expression/with_expression_chained_produces_independent_copies
// origin: languages/csharp/tests/csharp/test_csharp_with_expression.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Box(int Width, int Height);
var a = new Box(1, 1);
var b = a with { Width = 2 };
var c = b with { Height = 3 };
__Check((a.Width).ToString(), "1");
__Check((c.Width).ToString(), "2");
__Check((c.Height).ToString(), "3");
