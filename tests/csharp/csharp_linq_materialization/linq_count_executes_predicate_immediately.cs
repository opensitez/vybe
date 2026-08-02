// vybe-test: csharp/csharp_linq_materialization/linq_count_executes_predicate_immediately
// origin: languages/csharp/tests/csharp/test_csharp_linq_materialization.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Linq;
int checks = 0;
int total = new[] { 1, 2, 3 }.Count(x => { checks++; return x > 1; });
__Check((total).ToString(), "2");
__Check((checks).ToString(), "3");
