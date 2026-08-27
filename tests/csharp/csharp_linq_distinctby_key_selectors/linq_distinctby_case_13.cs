// vybe-test: csharp/csharp_linq_distinctby_key_selectors/linq_distinctby_case_13

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

var list = new (int Id, string Cat)[] { (1, "Tech"), (2, "Tech"), (3, "News") };
var distinct = list.DistinctBy(x => x.Cat).ToList();
__P(distinct.Count.ToString());
__P(distinct[0].Cat);
__P(distinct[1].Cat);
__Check("2\nTech\nNews");
