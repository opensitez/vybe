// vybe-test: csharp/csharp_delegate_variance/func_covariant_delegate_stored_in_object_field
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

System.Func<string> f=()=>"field"; System.Func<object> g=f; object boxed=g; __P((((System.Func<object>)boxed)()).ToString());
__Check("field");
