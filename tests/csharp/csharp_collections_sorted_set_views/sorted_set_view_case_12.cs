// vybe-test: csharp/csharp_collections_sorted_set_views/sorted_set_view_case_12

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

var ss = new System.Collections.Generic.SortedSet<int>() { 10, 20, 30, 40, 50 };
var view = ss.GetViewBetween(20, 40);
__P(view.Count.ToString());
__P(view.Min.ToString());
__P(view.Max.ToString());
__Check("3\n20\n40");
