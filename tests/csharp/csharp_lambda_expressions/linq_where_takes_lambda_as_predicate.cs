// vybe-test: csharp/csharp_lambda_expressions/linq_where_takes_lambda_as_predicate
// origin: languages/csharp/tests/csharp/test_csharp_lambda_expressions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var evens = new[]{1,2,3,4,5,6}.Where(n => n%2==0);
__Check((evens.Count()).ToString(), "3");
