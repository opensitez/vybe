// vybe-test: csharp/interfaces_generics/extension_method_on_list
// origin: languages/csharp/tests/csharp/test_interfaces_generics.rs

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

static class ListExtensions {
    public static string Join<T>(this List<T> list, string sep) {
        return string.Join(sep, list);
    }
}
var nums = new List<int> { 1, 2, 3, 4, 5 };
__P((nums.Join(", ")).ToString());
__Check("1, 2, 3, 4, 5");
