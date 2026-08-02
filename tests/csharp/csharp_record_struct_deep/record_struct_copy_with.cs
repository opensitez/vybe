// vybe-test: csharp/csharp_record_struct_deep/record_struct_copy_with
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct Count(int N); var a=new Count(5); var b=a; b=b with{N=99}; __Check((a.N).ToString(), "5"); __Check((b.N).ToString(), "99");
