// vybe-test: csharp/csharp_linq_aggregate_element/single_empty_throws_caught
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

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

string tag="ok";
try{System.Array.Empty<int>().Single();}catch(System.InvalidOperationException){tag="empty";}
__P((tag).ToString());
__Check("empty");
