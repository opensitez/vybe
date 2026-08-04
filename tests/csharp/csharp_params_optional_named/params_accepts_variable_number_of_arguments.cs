// vybe-test: csharp/csharp_params_optional_named/params_accepts_variable_number_of_arguments
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
__P((Sum(1,2,3,4,5)).ToString());
__Check("15");
