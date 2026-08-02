// vybe-test: csharp/csharp_deconstruct_tuples_records/tuple_swap
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int a=1,b=2; (a,b)=(b,a); __Check((a).ToString(), "2"); __Check((b).ToString(), "1");
