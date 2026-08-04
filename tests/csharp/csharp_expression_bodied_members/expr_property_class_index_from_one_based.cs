// vybe-test: csharp/csharp_expression_bodied_members/expr_property_class_index_from_one_based
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

class Row { public int Index = 0; public int Display => Index + 1; }
__P((new Row { Index = 4 }.Display).ToString());
__Check("5");
