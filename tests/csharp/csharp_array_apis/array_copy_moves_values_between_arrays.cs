// vybe-test: csharp/csharp_array_apis/array_copy_moves_values_between_arrays
// origin: languages/csharp/tests/csharp/test_csharp_array_apis.rs

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

var source = new[] { 5, 6, 7 }; var target = new int[3]; System.Array.Copy(source, target, 3); foreach (var value in target) __P((value).ToString());
__Check("5\n6\n7");
