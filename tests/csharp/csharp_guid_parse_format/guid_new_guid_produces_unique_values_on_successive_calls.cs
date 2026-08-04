// vybe-test: csharp/csharp_guid_parse_format/guid_new_guid_produces_unique_values_on_successive_calls
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_format.rs

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

var left = System.Guid.NewGuid();
var right = System.Guid.NewGuid();
__P((left != right).ToString());
__Check("True");
