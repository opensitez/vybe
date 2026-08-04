// vybe-test: csharp/csharp_linq_complex/aggregate_with_seed_computes_running_product
// origin: languages/csharp/tests/csharp/test_csharp_linq_complex.rs

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

var result=new[]{1,2,3,4,5}.Aggregate(1L,(acc,n)=>acc*n);
__P((result).ToString());
__Check("120");
