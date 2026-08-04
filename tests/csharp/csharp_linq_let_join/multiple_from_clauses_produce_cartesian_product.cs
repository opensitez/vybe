// vybe-test: csharp/csharp_linq_let_join/multiple_from_clauses_produce_cartesian_product
// origin: languages/csharp/tests/csharp/test_csharp_linq_let_join.rs

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

var result=from a in new[]{1,2} from b in new[]{10,20} select a*b;
int sum=0; foreach(var x in result) sum+=x;
__P((sum).ToString());
__Check("60");
