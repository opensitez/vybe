// vybe-test: csharp/common_patterns/binary_search_manual
// origin: languages/csharp/tests/csharp/test_common_patterns.rs

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

int[] arr = { 1, 3, 5, 7, 9, 11, 13 };
int target = 7;
int lo = 0, hi = arr.Length - 1;
while (lo <= hi) {
    int mid = (lo + hi) / 2;
    if (arr[mid] == target) { __P(("found at " + mid).ToString()); break; }
    else if (arr[mid] < target) lo = mid + 1;
    else hi = mid - 1;
}
__Check("found at 3");
