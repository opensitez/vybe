// vybe-test: csharp/csharp_with_expression_records/with_list_element
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record V(int N); var list=new System.Collections.Generic.List<V>{new V(1),new V(2)}; list[1]=list[1] with{N=9}; __Check((list[1].N).ToString(), "9");
