// vybe-test: csharp/csharp_expression_increment_semantics/increment_used_as_array_index_applies_after_index_is_read
// origin: languages/csharp/tests/csharp/test_csharp_expression_increment_semantics.rs

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

var data = new[] { 10, 20, 30 };
int i = 0;
__P((data[i++]).ToString());
__P((i).ToString());
__Check("10\n1");
