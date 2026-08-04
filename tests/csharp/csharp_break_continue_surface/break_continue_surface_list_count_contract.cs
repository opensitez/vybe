// vybe-test: csharp/csharp_break_continue_surface/break_continue_surface_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_break_continue_surface.rs

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

// break_continue_surface
var values = new System.Collections.Generic.List<int> { 49, 50, 49 }; __P((values.Count == 3).ToString());
__Check("True");
