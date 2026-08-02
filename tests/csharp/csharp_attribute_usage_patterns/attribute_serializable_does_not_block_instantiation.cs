// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_serializable_does_not_block_instantiation
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; [Serializable] class Node{public int Id=3;} __Check((new Node().Id).ToString(), "3");
