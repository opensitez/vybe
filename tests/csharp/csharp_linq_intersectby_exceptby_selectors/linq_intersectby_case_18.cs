// vybe-test: csharp/csharp_linq_intersectby_exceptby_selectors/linq_intersectby_case_18

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

var list1 = new (int Id, string Tag)[] { (1, "A"), (2, "B") };
string[] keys = new string[] { "A" };
var intersect = list1.IntersectBy(keys, x => x.Tag).ToList();
__P(intersect.Count.ToString());
__P(intersect[0].Id.ToString());
__Check("1\n1");
