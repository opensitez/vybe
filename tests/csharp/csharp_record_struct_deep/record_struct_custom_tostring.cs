// vybe-test: csharp/csharp_record_struct_deep/record_struct_custom_tostring
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Tag(string Name){public override string ToString()=>"Tag:"+Name;} __Check((new Tag("x")).ToString(), "Tag:x");
