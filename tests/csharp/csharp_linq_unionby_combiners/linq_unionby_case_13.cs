// vybe-test: csharp/csharp_linq_unionby_combiners/linq_unionby_case_13

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

var l1 = new (int Id, string N)[] { (1, "A"), (2, "B") };
var l2 = new (int Id, string N)[] { (2, "B_Dupe"), (3, "C") };
var union = l1.UnionBy(l2, x => x.Id).ToList();
__P(union.Count.ToString());
__Check("3");
