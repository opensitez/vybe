// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_dictionary_values_to_list_inferred
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

System.Collections.Generic.Dictionary<string, int> map = new() { ["a"] = 1, ["b"] = 2 };
System.Collections.Generic.List<int> values = new();
foreach (var kv in map) values.Add(kv.Value);
__P((values[0] + values[1]).ToString());
__Check("3");
