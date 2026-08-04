// vybe-test: csharp/csharp_method_overloading/overload_on_type_picks_most_specific_match
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

string Label(object o)=>"object";
string Label(string s)=>"string";
__P((Label("hi")).ToString());
__P((Label((object)"hi")).ToString());
__Check("string\nobject");
