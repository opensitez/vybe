// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_serializable_array_of_instances
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Serializable] class Node{public int Id;} var arr=new Node[]{new Node{Id=1},new Node{Id=2}}; __Check((arr[1].Id).ToString(), "2");
