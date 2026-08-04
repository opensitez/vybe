// vybe-test: csharp/csharp_attribute_usage_patterns/attribute_serializable_does_not_block_instantiation
// origin: languages/csharp/tests/csharp/test_csharp_attribute_usage_patterns.rs

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

using System; [Serializable] class Node{public int Id=3;} __P((new Node().Id).ToString());
__Check("3");
