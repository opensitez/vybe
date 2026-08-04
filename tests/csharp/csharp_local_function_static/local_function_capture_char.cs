// vybe-test: csharp/csharp_local_function_static/local_function_capture_char
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

char ch='A'; string Show(int n){string S(int x)=>ch+""+x; return S(n);} __P((Show(1)).ToString());
__Check("A1");
