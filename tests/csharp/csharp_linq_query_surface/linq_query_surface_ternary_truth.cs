// vybe-test: csharp/csharp_linq_query_surface/linq_query_surface_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_linq_query_surface.rs

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

// linq_query_surface
int seed = 117; bool cond = seed % 2 == 0; __P((cond || !cond).ToString());
__Check("True");
