// vybe-test: csharp/csharp_null_handling/null_conditional_chains_through_nested_properties
// origin: languages/csharp/tests/csharp/test_csharp_null_handling.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Node { public Node Next; public int Value; }
Node head = null;
__Check((head?.Next?.Value ?? -1).ToString(), "-1");
