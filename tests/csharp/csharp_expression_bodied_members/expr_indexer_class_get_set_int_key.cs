// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_get_set_int_key
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Buffer { int[] data = new int[3]; public int this[int i] { get => data[i]; set => data[i] = value; } }
var b = new Buffer(); b[2] = 99; __Check((b[2]).ToString(), "99");
