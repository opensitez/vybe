// vybe-test: csharp/csharp_local_functions_partial_methods/partial_method_can_be_triggered_from_constructor
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

partial class Worker { partial void OnCreated(); public Worker() { OnCreated(); } } partial class Worker { partial void OnCreated() { System.__P(("created").ToString()); } } new Worker();
__Check("created");
