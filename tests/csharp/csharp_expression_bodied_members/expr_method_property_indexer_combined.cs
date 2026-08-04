// vybe-test: csharp/csharp_expression_bodied_members/expr_method_property_indexer_combined
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

class Cache { int[] buf = { 0, 0, 0 }; public int this[int i] { get => buf[i]; set => buf[i] = value; } public int Sum() => buf[0] + buf[1] + buf[2]; }
var c = new Cache(); c[0] = 1; c[1] = 2; c[2] = 3; __P((c.Sum()).ToString());
__Check("6");
