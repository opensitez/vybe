// vybe-test: csharp/csharp_generic_variance2/covariant_interface_allows_derived_where_base_expected
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

interface IReader<out T>{T Read();}
class StringReader:IReader<string>{public string Read()=>"hello";}
IReader<object> r=new StringReader();
__P((r.Read()).ToString());
__Check("hello");
