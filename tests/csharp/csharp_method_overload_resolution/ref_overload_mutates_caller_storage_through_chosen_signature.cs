// vybe-test: csharp/csharp_method_overload_resolution/ref_overload_mutates_caller_storage_through_chosen_signature
// origin: languages/csharp/tests/csharp/test_csharp_method_overload_resolution.rs

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

void Scale(int value) { __P(("byval:" + value).ToString()); }
void Scale(ref int value) { value = value * 2; }
int n = 5;
Scale(ref n);
__P(("after:" + n).ToString());
__Check("after:10");
