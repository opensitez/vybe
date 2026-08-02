// vybe-test: csharp/csharp_ref_readonly_semantics/readonly_ref_struct_inequality_when_ids_differ
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

readonly ref struct Tag{public readonly int Id; public Tag(int id){Id=id;} public bool Equals(Tag other)=>Id==other.Id;} var a=new Tag(1); var b=new Tag(2); __Check((a.Equals(b)).ToString(), "False");
