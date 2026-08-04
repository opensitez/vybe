// vybe-test: csharp/csharp_covariance_contravariance/func_return_type_covariance_allows_derived_func_in_base_func
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

System.Func<string> getString = () => "hi";
System.Func<object> getObject = getString;
__P((getObject()).ToString());
__Check("hi");
