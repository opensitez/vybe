// vybe-test: csharp/csharp_collections_immutable_hashset_algebra/immutable_hashset_case_9

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var s1 = System.Collections.Immutable.ImmutableHashSet.Create<int>(1, 2, 3);
var s2 = System.Collections.Immutable.ImmutableHashSet.Create<int>(3, 4, 5);
var union = s1.Union(s2);
__P(union.Count.ToString());
__Check("5");
