// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_void_writes_argument
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Echo { public void Say(string msg) => __Check((msg).ToString(), "hi"); }
new Echo().Say("hi");
