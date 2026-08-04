// vybe-test: csharp/csharp_expression_bodied_members/expr_indexer_class_get_set_int_key
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

class Buffer { int[] data = new int[3]; public int this[int i] { get => data[i]; set => data[i] = value; } }
var b = new Buffer(); b[2] = 99; __P((b[2]).ToString());
__Check("99");
