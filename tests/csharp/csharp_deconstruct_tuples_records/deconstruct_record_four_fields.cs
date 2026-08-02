// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_record_four_fields
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Quad(int A,int B,int C,int D); var (a,b,c,d)=new Quad(1,2,3,4); __Check((d).ToString(), "4");
