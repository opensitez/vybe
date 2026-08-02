// vybe-test: csharp/csharp_hashset_set_algebra/set_equals_with_enumerable_matching_set
// origin: languages/csharp/tests/csharp/test_csharp_hashset_set_algebra.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var a = new HashSet<int> { 4, 5 }; __Check((a.SetEquals(new[] { 5, 4 })).ToString(), "True");
