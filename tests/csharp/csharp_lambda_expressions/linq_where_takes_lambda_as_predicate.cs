// vybe-test: csharp/csharp_lambda_expressions/linq_where_takes_lambda_as_predicate
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

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

var evens = new[]{1,2,3,4,5,6}.Where(n => n%2==0);
__P((evens.Count()).ToString());
__Check("3");
