// vybe-test: csharp/csharp_delegates_advanced/lambda_closed_over_mutable_list_builds_result
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

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

var log=new System.Collections.Generic.List<string>();
System.Action<string> record=msg=>log.Add(msg);
record("a"); record("b"); record("c");
__P((string.Join(",",log)).ToString());
__Check("a,b,c");
