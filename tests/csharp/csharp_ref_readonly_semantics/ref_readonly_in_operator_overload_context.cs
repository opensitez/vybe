// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_in_operator_overload_context
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly struct Num{public readonly int Value; public Num(int v){Value=v;} public static bool operator ==(Num a, ref readonly Num b)=>a.Value==b.Value; public static bool operator !=(Num a, ref readonly Num b)=>!(a==b);} var x=new Num(4); var y=new Num(4); __Check((x==ref y).ToString(), "True");
