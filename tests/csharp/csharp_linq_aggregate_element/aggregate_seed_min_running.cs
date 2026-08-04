// vybe-test: csharp/csharp_linq_aggregate_element/aggregate_seed_min_running
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

var min=new[]{3,1,4,1,5}.Aggregate(int.MaxValue,(acc,x)=>x<acc?x:acc);
__P((min).ToString());
__Check("1");
