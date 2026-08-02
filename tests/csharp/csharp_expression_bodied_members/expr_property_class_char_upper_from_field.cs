// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_char_upper_from_field
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Token { public char ch = 'a'; public char Upper => char.ToUpper(ch); }
__Check((new Token().Upper).ToString(), "A");
