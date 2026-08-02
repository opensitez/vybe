// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_two_params_sums
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Adder { public int Sum(int a, int b) => a + b; }
__Check((new Adder().Sum(3, 4)).ToString(), "7");
