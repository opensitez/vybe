// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_field_in_readonly_struct
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly struct Pair{public readonly int First; public readonly int Second; public Pair(int a,int b){First=a; Second=b;}} var p=new Pair(2,3); __Check((p.First+p.Second).ToString(), "5");
