// vybe-test: csharp/csharp_linq_aggregate_element/single_with_predicate_many_matches_caught
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
try{new[]{1,2,2}.Single(x=>x==2);}catch(System.InvalidOperationException){tag="many";}
__P((tag).ToString());
__Check("many");
