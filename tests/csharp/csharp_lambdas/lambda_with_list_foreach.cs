// vybe-test: csharp/csharp_lambdas/lambda_with_list_foreach
// origin: languages/csharp/tests/csharp/test_csharp_lambdas.rs

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

using System.Collections.Generic;
var items = new List<int>();
items.Add(1);
items.Add(2);
items.Add(3);
items.ForEach(x => __P((x).ToString()));
__Check("1\n2\n3");
