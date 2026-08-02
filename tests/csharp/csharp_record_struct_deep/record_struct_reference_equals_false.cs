// vybe-test: csharp/csharp_record_struct_deep/record_struct_reference_equals_false
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Key(int Id); var a=new Key(1); var b=new Key(1); __Check((System.Object.ReferenceEquals(a,b)).ToString(), "False");
