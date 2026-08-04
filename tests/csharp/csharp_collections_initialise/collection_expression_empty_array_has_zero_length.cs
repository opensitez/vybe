// vybe-test: csharp/csharp_collections_initialise/collection_expression_empty_array_has_zero_length
// origin: languages/csharp/tests/csharp/test_csharp_collections_initialise.rs

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

int[] empty=[];
__P((empty.Length).ToString());
__Check("0");
