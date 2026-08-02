// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_returns_char_from_index
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pick { public char At(string s, int i) => s[i]; }
__Check((new Pick().At("cat", 1)).ToString(), "a");
