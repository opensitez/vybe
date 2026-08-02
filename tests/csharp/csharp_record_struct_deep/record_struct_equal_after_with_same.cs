// vybe-test: csharp/csharp_record_struct_deep/record_struct_equal_after_with_same
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct V(int N); var a=new V(1); var b=a with{N=1}; __Check((a==b).ToString(), "True");
