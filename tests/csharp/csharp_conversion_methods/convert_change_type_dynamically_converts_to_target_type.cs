// vybe-test: csharp/csharp_conversion_methods/convert_change_type_dynamically_converts_to_target_type
// origin: languages/csharp/tests/csharp/test_csharp_conversion_methods.rs

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

object result=System.Convert.ChangeType("42",typeof(int));
__P((result).ToString());
__Check("42");
