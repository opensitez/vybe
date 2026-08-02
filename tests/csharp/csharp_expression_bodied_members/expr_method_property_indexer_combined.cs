// vybe-test: csharp/csharp_expression_bodied_members/expr_method_property_indexer_combined
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Cache { int[] buf = { 0, 0, 0 }; public int this[int i] { get => buf[i]; set => buf[i] = value; } public int Sum() => buf[0] + buf[1] + buf[2]; }
var c = new Cache(); c[0] = 1; c[1] = 2; c[2] = 3; __Check((c.Sum()).ToString(), "6");
