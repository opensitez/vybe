// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_get_only_int_key
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

class Bag { int[] data = { 10, 20, 30 }; public int this[int i] => data[i]; }
__P((new Bag()[1]).ToString());
__Check("20");
