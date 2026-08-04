// vybe-test: csharp/csharp_deferred_execution/count_vs_to_list_count_return_same_number
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

var q=new[]{1,2,3,4}.Where(x=>x%2==0);
__P((q.Count()).ToString());
__P((q.ToList().Count).ToString());
__Check("2\n2");
