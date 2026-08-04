// vybe-test: csharp/csharp_local_function_static/local_function_captures_local_struct
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

int UseStruct(){var p=new System.ValueTuple<int,int>(2,3); int Sum(int n){int S(int x)=>p.Item1+p.Item2+x; return S(n);} return Sum(1);} __P((UseStruct()).ToString());
__Check("6");
