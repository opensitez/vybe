// vybe-test: csharp/csharp_local_function_static/static_local_function_called_from_sibling_local
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

int Pipeline(int n){static int Double(int x)=>x*2; int Wrap(int v)=>Double(v)+1; return Wrap(n);} __P((Pipeline(5)).ToString());
__Check("11");
