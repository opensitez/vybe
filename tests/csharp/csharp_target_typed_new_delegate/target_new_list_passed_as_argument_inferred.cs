// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_passed_as_argument_inferred
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

int Sum(System.Collections.Generic.List<int> xs) { int s = 0; foreach (var x in xs) s += x; return s; }
System.Collections.Generic.List<int> data = new() { 1, 2, 3 };
__P((Sum(data)).ToString());
__Check("6");
