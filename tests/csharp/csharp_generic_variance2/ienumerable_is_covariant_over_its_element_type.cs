// vybe-test: csharp/csharp_generic_variance2/ienumerable_is_covariant_over_its_element_type
// origin: languages/csharp/tests/csharp/test_csharp_generic_variance2.rs

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

System.Collections.Generic.IEnumerable<string> strings=new[]{"a","b"};
System.Collections.Generic.IEnumerable<object> objects=strings;
__P((objects.Count()).ToString());
__Check("2");
