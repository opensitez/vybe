// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

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

// char_predicate_apis
var values = new System.Collections.Generic.List<int> { 23, 24, 23 }; __P((values.Count == 3).ToString());
__Check("True");
