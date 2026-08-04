// vybe-test: csharp/csharp_collections_initialise/span_collection_expression_works_with_stack_alloc_semantics
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

System.Span<int> s=[1,2,3];
__P((s.Length).ToString());
__Check("3");
