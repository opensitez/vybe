// vybe-test: csharp/csharp_deferred_execution/any_short_circuits_after_first_match
// origin: languages/csharp/tests/csharp/test_csharp_deferred_execution.rs

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

int count=0;
bool found=new[]{1,2,3,4,5}.Any(n=>{count++;return n==3;});
__P((found).ToString()); __P((count).ToString());
__Check("True\n3");
