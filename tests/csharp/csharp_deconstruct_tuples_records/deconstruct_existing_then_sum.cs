// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_existing_then_sum
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int s=0,t=0; (s,t)=(4,6); __Check((s+t).ToString(), "10");
