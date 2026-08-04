// vybe-test: csharp/csharp_list_operations/to_array_converts_list_to_fixed_array
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

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

var list = new System.Collections.Generic.List<int>{7,8,9};
var arr = list.ToArray();
__P((arr.GetType().IsArray).ToString());
__Check("True");
