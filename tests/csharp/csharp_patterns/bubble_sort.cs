// vybe-test: csharp/csharp_patterns/bubble_sort
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

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

var arr = new[] { 5, 3, 8, 1, 2 };
for (int i = 0; i < arr.Length - 1; i++) {
    for (int j = 0; j < arr.Length - 1 - i; j++) {
        if (arr[j] > arr[j + 1]) {
            int temp = arr[j];
            arr[j] = arr[j + 1];
            arr[j + 1] = temp;
        }
    }
}
foreach (var x in arr) __P((x).ToString());
__Check("1\n2\n3\n5\n8");
