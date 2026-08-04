// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_three_params_product
// origin: languages/csharp/tests/csharp/test_csharp_expression_bodied_members.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Mul3 { public int Prod(int a, int b, int c) => a * b * c; }
__P((new Mul3().Prod(2, 3, 4)).ToString());
__Check("24");
