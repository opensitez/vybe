// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_chained_computed
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

class Chain { public int Base = 2; public int Double => Base * 2; public int Quadruple => Double * 2; }
__P((new Chain().Quadruple).ToString());
__Check("8");
