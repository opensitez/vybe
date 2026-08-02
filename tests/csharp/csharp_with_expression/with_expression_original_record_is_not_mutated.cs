// vybe-test: csharp/csharp_with_expression/with_expression_original_record_is_not_mutated
// origin: languages/csharp/tests/csharp/test_csharp_with_expression.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Point(int X, int Y);
var origin = new Point(1, 2);
var moved = origin with { X = 10 };
__Check((origin.X).ToString(), "1");
