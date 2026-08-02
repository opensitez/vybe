// vybe-test: csharp/csharp_ref_readonly_semantics/readonly_ref_struct_can_be_local_variable
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly ref struct Marker{public readonly int Code; public Marker(int c){Code=c;}} var m=new Marker(42); __Check((m.Code).ToString(), "42");
