// vybe-test: csharp/csharp_deconstruct_tuples_records/custom_deconstruct_discard
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pair{public int A,B; public void Deconstruct(out int a,out int b){a=A;b=B;}} var (_,b)=new Pair{A=9,B=2}; __Check((b).ToString(), "2");
