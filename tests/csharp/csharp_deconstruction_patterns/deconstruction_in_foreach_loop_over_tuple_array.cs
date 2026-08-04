// vybe-test: csharp/csharp_deconstruction_patterns/deconstruction_in_foreach_loop_over_tuple_array
// origin: languages/csharp/tests/csharp/test_csharp_deconstruction_patterns.rs

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

var pairs = new[]{(1,"a"),(2,"b"),(3,"c")};
int sum=0;
foreach(var (n, _) in pairs) sum+=n;
__P((sum).ToString());
__Check("6");
