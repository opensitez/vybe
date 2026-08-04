// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_then_element_at
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

var running=new[]{1,2,3}.Aggregate(new int[]{0},(acc,x)=>new int[]{acc[0]+x});
__P((running[0]).ToString());
__Check("6");
