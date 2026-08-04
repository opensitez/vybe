// vybe-test: csharp/csharp_new_features/target_typed_new_infers_list_type_from_variable
// origin: languages/csharp/tests/csharp/test_csharp_new_features.rs

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

System.Collections.Generic.List<int> nums = new();
nums.Add(1); nums.Add(2);
__P((nums.Count).ToString());
__Check("2");
