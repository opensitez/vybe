// vybe-test: csharp/csharp_list_dictionary/dictionary_int_key_overwrite_via_indexer
// origin: languages/csharp/tests/csharp/test_csharp_list_dictionary.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using System.Collections.Generic; var map = new Dictionary<int, int> { [1] = 10 }; map[1] = 99; __P((map[1]).ToString());
__Check("99");
