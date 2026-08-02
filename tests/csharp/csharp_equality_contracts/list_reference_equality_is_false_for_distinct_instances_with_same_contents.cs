// vybe-test: csharp/csharp_equality_contracts/list_reference_equality_is_false_for_distinct_instances_with_same_contents
// origin: languages/csharp/tests/csharp/test_csharp_equality_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic;
using System.Linq;
var left = new List<int> { 1, 2 };
var right = new List<int> { 1, 2 };
__Check((left == right).ToString(), "False");
__Check((left.SequenceEqual(right)).ToString(), "True");
