// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_hashset_string_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

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

System.Collections.Generic.HashSet<string> tags = new();
tags.Add("a"); tags.Add("b");
__P((tags.Contains("a")).ToString()); __P((tags.Contains("c")).ToString());
__Check("True\nFalse");
