// vybe-test: csharp/csharp_covariance_contravariance/contravariant_in_allows_base_generic_argument_on_interface
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

interface IWriter<in T> { void Write(T value); }
class ObjectWriter : IWriter<object> {
    public void Write(object value) => __P((value).ToString());
}
IWriter<string> writer = new ObjectWriter();
writer.Write("typed");
__Check("typed");
