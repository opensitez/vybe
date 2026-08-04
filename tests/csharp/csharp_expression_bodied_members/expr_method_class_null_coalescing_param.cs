// vybe-test: csharp/csharp_expression_bodied_members/expr_method_class_null_coalescing_param
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

class Safe { public string OrEmpty(string? s) => s ?? ""; }
__P((new Safe().OrEmpty(null)).ToString()); __P((new Safe().OrEmpty("x")).ToString());
__Check("\nx");
