// vybe-test: csharp/csharp_params_optional_named/named_argument_can_be_passed_out_of_order
// origin: languages/csharp/tests/csharp/test_csharp_params_optional_named.rs

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

string Concat(string a, string b, string c) => a+b+c;
__P((Concat(c:"3",a:"1",b:"2")).ToString());
__Check("123");
