// vybe-test: csharp/csharp_local_functions/local_function_used_before_its_declaration
// origin: languages/csharp/tests/csharp/test_csharp_local_functions.rs

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

__P((Double(5)).ToString());
int Double(int x)=>x*2;
__Check("10");
