// vybe-test: csharp/csharp_extension_methods_patterns/extension_method_on_ienumerable_implements_custom_linq_step
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods_patterns.rs

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

static class SeqExt {
    public static System.Collections.Generic.IEnumerable<T> Every<T>(
        this System.Collections.Generic.IEnumerable<T> src, int n) {
        int i=0; foreach(var x in src) { if(i++%n==0) yield return x; }
    }
}
var result = new[]{1,2,3,4,5,6}.Every(2);
foreach(var x in result) __P((x).ToString());
__Check("1\n3\n5");
