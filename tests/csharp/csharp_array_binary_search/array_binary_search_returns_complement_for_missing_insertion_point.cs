// vybe-test: csharp/csharp_array_binary_search/array_binary_search_returns_complement_for_missing_insertion_point
// origin: languages/csharp/tests/csharp/test_csharp_array_binary_search.rs

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

int[] sorted = { 2, 4, 8 };
int index = System.Array.BinarySearch(sorted, 5);
__P((index < 0).ToString());
__P((~index).ToString());
__Check("True\n2");
