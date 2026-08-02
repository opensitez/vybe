// vybe-test: csharp/csharp_deconstruct_tuples_records/deconstruct_to_existing_locals
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pair{public int A,B; public void Deconstruct(out int a,out int b){a=A;b=B;}} var target=new Pair{A=5,B=6}; int x,y; (x,y)=target; __Check((x+y).ToString(), "11");
