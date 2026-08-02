// vybe-test: csharp/csharp_lambda_expressions/lambda_captures_and_modifies_outer_list
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var results = new System.Collections.Generic.List<int>();
var nums = new[]{1,2,3,4};
System.Array.ForEach(nums, n => { if(n%2==0) results.Add(n); });
__Check((results.Count).ToString(), "2");
