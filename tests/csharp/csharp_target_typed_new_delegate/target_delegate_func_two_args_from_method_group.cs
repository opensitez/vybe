// vybe-test: csharp/csharp_target_typed_new_delegate/target_delegate_func_two_args_from_method_group
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

static int Add(int a, int b) => a + b;
System.Func<int, int, int> add = Add;
__P((add(3, 4)).ToString());
__Check("7");
