// vybe-test: csharp/csharp_oop_polymorphism/direct_cast_throws_invalid_cast_for_unrelated_type
// origin: languages/csharp/tests/csharp/test_csharp_oop_polymorphism.rs

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

string r="";
try{object o="hello"; int n=(int)o;}
catch(System.InvalidCastException){r="bad cast";}
__P((r).ToString());
__Check("bad cast");
