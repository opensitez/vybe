// vybe-test: csharp/csharp_delegate_types/func_t_result_returns_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_delegate_types.rs

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

System.Func<int,int> square = x => x * x;
__P((square(4)).ToString());
__Check("16");
