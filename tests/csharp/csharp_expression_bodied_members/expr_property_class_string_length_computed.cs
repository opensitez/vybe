// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_string_length_computed
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Label { public string Text = "hello"; public int Len => Text.Length; }
__Check((new Label().Len).ToString(), "5");
