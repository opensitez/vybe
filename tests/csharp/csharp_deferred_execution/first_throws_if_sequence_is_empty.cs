// vybe-test: csharp/csharp_deferred_execution/first_throws_if_sequence_is_empty
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

string r="";
try{System.Array.Empty<int>().First();}
catch(System.InvalidOperationException){r="empty";}
__P((r).ToString());
__Check("empty");
