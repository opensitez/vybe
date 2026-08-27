// vybe-test: csharp/csharp_linq_minby_maxby_selectors/linq_minby_maxby_case_20

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

var list = new (string Name, int Val)[] { ("A", 10), ("B", 20), ("C", 5) };
var min = list.MinBy(x => x.Val);
var max = list.MaxBy(x => x.Val);
__P(min.Name);
__P(max.Name);
__Check("C\nB");
