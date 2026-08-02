// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_double_record
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Rate(double V); var (v)=new Rate(3.5); __Check((v).ToString(), "3.5");
