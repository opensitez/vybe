// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_two_int_params
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

class Grid { int[,] m = { { 1, 2 }, { 3, 4 } }; public int this[int r, int c] => m[r, c]; }
__P((new Grid()[1, 0]).ToString());
__Check("3");
