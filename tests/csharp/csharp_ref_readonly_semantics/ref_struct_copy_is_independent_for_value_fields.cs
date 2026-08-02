// vybe-test: csharp/csharp_ref_readonly_semantics/ref_struct_copy_is_independent_for_value_fields
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

ref struct Box{public int Item;} var x=new Box(); x.Item=10; var y=x; y.Item=99; __Check((x.Item).ToString(), "10");
