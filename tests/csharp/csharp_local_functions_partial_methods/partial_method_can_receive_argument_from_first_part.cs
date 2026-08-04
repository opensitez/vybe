// vybe-test: csharp/csharp_local_functions_partial_methods/partial_method_can_receive_argument_from_first_part
// origin: languages/csharp/tests/csharp/test_csharp_local_functions_partial_methods.rs

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

partial class Worker { partial void OnRun(int value); public void Run() { OnRun(5); } } partial class Worker { partial void OnRun(int value) { System.__P((value * 2).ToString()); } } new Worker().Run();
__Check("10");
