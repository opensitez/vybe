// vybe-test: csharp/csharp_array_reverse_patterns/array_reverse_patterns_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_array_reverse_patterns.rs

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

// array_reverse_patterns
int seed = 27; int right = seed + 1; __P((seed < right).ToString());
__Check("True");
