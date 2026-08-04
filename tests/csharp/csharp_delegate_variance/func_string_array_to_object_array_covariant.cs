// vybe-test: csharp/csharp_delegate_variance/func_string_array_to_object_array_covariant
// origin: languages/csharp/tests/csharp/test_csharp_delegate_variance.rs

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

System.Func<string[]> getStrings=()=>new string[]{"a"}; System.Func<object[]> getObjects=getStrings; __P((getObjects()[0]).ToString());
__Check("a");
