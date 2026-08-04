// vybe-test: csharp/csharp_queue_stack_surface/queue_stack_surface_string_contains_probe
// origin: languages/csharp/tests/csharp/test_csharp_queue_stack_surface.rs

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

// queue_stack_surface
string feature = "queue_stack_surface"; __P((feature.Contains("a") || !feature.Contains("a")).ToString());
__Check("True");
