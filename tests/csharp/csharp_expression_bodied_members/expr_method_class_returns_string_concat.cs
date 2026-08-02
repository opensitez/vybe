// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_returns_string_concat
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Joiner { public string Merge(string a, string b) => a + b; }
__Check((new Joiner().Merge("ab", "cd")).ToString(), "abcd");
