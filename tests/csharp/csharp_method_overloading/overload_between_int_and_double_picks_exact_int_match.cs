// vybe-test: csharp/csharp_method_overloading/overload_between_int_and_double_picks_exact_int_match
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

string Kind(int n)=>"int";
string Kind(double d)=>"double";
__P((Kind(5)).ToString());
__P((Kind(5.0)).ToString());
__Check("int\ndouble");
