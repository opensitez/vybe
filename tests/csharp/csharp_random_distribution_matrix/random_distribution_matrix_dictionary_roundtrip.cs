// vybe-test: csharp/csharp_random_distribution_matrix/random_distribution_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_random_distribution_matrix.rs

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

// random_distribution_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[98] = 99; __P((map.ContainsKey(98) && map[98] == 99).ToString());
__Check("True");
