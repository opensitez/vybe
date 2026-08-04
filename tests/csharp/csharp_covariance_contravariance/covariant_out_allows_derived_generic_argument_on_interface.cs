// vybe-test: csharp/csharp_covariance_contravariance/covariant_out_allows_derived_generic_argument_on_interface
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

interface IReader<out T> { T Read(); }
class StringReader : IReader<string> {
    public string Read() => "hello";
}
IReader<object> reader = new StringReader();
__P((reader.Read()).ToString());
__Check("hello");
