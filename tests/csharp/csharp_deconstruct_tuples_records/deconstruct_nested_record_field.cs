// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_nested_record_field
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Inner(int N); record Outer(Inner I); var (n)=new Outer(new Inner(9)); __Check((n).ToString(), "9");
