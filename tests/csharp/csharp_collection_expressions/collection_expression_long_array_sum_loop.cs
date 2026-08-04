// vybe-test: csharp/csharp_collection_expressions/collection_expression_long_array_sum_loop
// origin: languages/csharp/tests/csharp/test_csharp_collection_expressions.rs

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

long[] nums = [10000000000L, 20000000000L];
long total = 0;
foreach (var n in nums) total += n;
__P((total).ToString());
__Check("30000000000");
