// vybe-test: csharp/csharp_method_overloading/overload_with_params_array_chosen_when_explicit_available
// origin: languages/csharp/tests/csharp/test_csharp_method_overloading.rs

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

string Sum(int a,int b)=>"two";
string Sum(params int[] ns)=>"params";
__P((Sum(1,2)).ToString());
__P((Sum(1,2,3)).ToString());
__Check("two\nparams");
