// vybe-test: csharp/csharp_collection_expressions/collection_expression_bool_array_values
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

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

bool[] flags = [true, false, true];
__P((flags[0]).ToString()); __P((flags[1]).ToString());
__Check("True\nFalse");
