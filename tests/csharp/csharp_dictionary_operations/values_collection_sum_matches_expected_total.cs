// vybe-test: csharp/csharp_dictionary_operations/values_collection_sum_matches_expected_total
// origin: languages/csharp/tests/csharp/test_csharp_dictionary_operations.rs

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

var d = new System.Collections.Generic.Dictionary<string,int>{{"a",3},{"b",7}};
int sum=0; foreach(var v in d.Values) sum+=v;
__P((sum).ToString());
__Check("10");
