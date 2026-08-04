// vybe-test: csharp/csharp_local_function_static/local_function_bool_predicate
// origin: languages/csharp/tests/csharp/test_csharp_local_function_static.rs

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

bool AllPositive(int a,int b){bool Check(int x,int y)=>x>0&&y>0; return Check(a,b);} __P((AllPositive(1,2)).ToString());
__Check("True");
