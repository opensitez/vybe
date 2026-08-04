// vybe-test: csharp/csharp_params_optional_named/params_can_be_called_with_explicit_array
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

int Sum(params int[] ns){int s=0;foreach(var n in ns)s+=n;return s;}
__P((Sum(new int[]{10,20})).ToString());
__Check("30");
