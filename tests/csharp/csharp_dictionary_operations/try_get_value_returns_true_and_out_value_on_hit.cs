// vybe-test: csharp/csharp_dictionary_operations/try_get_value_returns_true_and_out_value_on_hit
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

var d = new System.Collections.Generic.Dictionary<string,int>{{"k",5}};
__P((d.TryGetValue("k", out int v)).ToString());
__P((v).ToString());
__Check("True\n5");
