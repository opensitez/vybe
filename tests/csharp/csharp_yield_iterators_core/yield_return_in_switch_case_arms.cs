// vybe-test: csharp/csharp_yield_iterators_core/yield_return_in_switch_case_arms
// origin: languages/csharp/tests/csharp/test_csharp_yield_iterators_core.rs

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

System.Collections.Generic.IEnumerable<string> Label(int n){switch(n){case 1:yield return "one";break;case 2:yield return "two";break;default:yield return "many";break;}}
__P((string.Join("|",Label(2))).ToString());
__Check("two");
