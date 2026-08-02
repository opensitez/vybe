// vybe-test: csharp/csharp_expression_bodied_members/expr_method_expression_body_with_conditional
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Sign { public string Label(int n) => n < 0 ? "neg" : n > 0 ? "pos" : "zero"; }
__Check((new Sign().Label(-1)).ToString(), "neg"); __Check((new Sign().Label(0)).ToString(), "zero"); __Check((new Sign().Label(2)).ToString(), "pos");
