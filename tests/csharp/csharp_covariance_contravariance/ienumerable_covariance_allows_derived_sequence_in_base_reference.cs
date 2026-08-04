// vybe-test: csharp/csharp_covariance_contravariance/ienumerable_covariance_allows_derived_sequence_in_base_reference
// origin: languages/csharp/tests/csharp/test_csharp_covariance_contravariance.rs

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

System.Collections.Generic.IEnumerable<string> strings =
    new System.Collections.Generic.List<string> { "x" };
System.Collections.Generic.IEnumerable<object> objects = strings;
foreach (var o in objects) __P((o).ToString());
__Check("x");
