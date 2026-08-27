// vybe-test: csharp/csharp_collections_immutable_array_mutations/immutable_array_mutation_case_2

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

var a1 = System.Collections.Immutable.ImmutableArray.Create<int>(2);
var a2 = a1.Add(3);
__P(a1.Length.ToString());
__P(a2.Length.ToString());
__P(a2[1].ToString());
__Check("1\n2\n3");
