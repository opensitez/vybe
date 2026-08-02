// vybe-test: csharp/csharp_record_struct_deep/record_struct_array_index
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct V(int N); var arr=new[]{new V(1),new V(2)}; __Check((arr[1].N).ToString(), "2");
