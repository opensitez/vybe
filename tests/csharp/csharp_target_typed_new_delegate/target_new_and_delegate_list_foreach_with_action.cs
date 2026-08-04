// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_and_delegate_list_foreach_with_action
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

System.Collections.Generic.List<int> nums = new() { 1, 2, 3 };
int sum = 0;
System.Action<int> acc = n => sum += n;
foreach (var n in nums) acc(n);
__P((sum).ToString());
__Check("6");
