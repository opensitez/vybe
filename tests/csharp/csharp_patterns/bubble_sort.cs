// vybe-test: csharp/csharp_patterns/bubble_sort
// origin: languages/csharp/tests/csharp/test_csharp_patterns.rs

using static __Harness;

var arr = new[] { 5, 3, 8, 1, 2 }
;
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

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
