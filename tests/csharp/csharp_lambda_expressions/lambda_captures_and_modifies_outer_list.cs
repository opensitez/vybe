// vybe-test: csharp/csharp_lambda_expressions/lambda_captures_and_modifies_outer_list
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

var results = new System.Collections.Generic.List<int>();
var nums = new[]{1,2,3,4};
System.Array.ForEach(nums, n => { if(n%2==0) results.Add(n); });
__P((results.Count).ToString());
__Check("2");
