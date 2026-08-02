// vybe-test: csharp/csharp_record_struct_deep/record_struct_equals_object_equal
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Key(int Id); object o=new Key(5); __Check((new Key(5).Equals(o)).ToString(), "True");
