// vybe-test: csharp/csharp_deconstruct_tuples_records/custom_class_deconstruct_single
// origin: languages/csharp/tests/csharp/test_csharp_deconstruct_tuples_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Wrap{public int V; public void Deconstruct(out int v){v=V;}} var (v)=new Wrap{V=11}; __Check((v).ToString(), "11");
