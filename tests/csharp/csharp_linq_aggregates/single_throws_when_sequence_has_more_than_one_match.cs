// vybe-test: csharp/csharp_linq_aggregates/single_throws_when_sequence_has_more_than_one_match
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregates.rs

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

string result = "ok";
try { new[]{1,2}.Single(); }
catch(System.InvalidOperationException) { result = "many"; }
__P((result).ToString());
__Check("many");
