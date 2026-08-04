// vybe-test: csharp/csharp_local_functions_partial_methods/partial_method_implemented_in_second_part_can_be_called
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

partial class Worker { partial void OnRun(); public void Run() { OnRun(); } } partial class Worker { partial void OnRun() { System.__P(("ran").ToString()); } } new Worker().Run();
__Check("ran");
