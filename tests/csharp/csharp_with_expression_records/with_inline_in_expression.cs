// vybe-test: csharp/csharp_with_expression_records/with_inline_in_expression
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record V(int N); __Check(((new V(2) with{N=7}).N).ToString(), "7");
