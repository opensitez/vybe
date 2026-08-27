// vybe-test: csharp/csharp_concurrent_bag_unordered_work/concurrent_bag_work_case_19

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

var bag = new System.Collections.Concurrent.ConcurrentBag<string>();
bag.Add("Item_19");
bool pk = bag.TryPeek(out string p);
bool tk = bag.TryTake(out string t);
__P(pk.ToString());
__P(p);
__P(tk.ToString());
__P(t);
__Check("True\nItem_19\nTrue\nItem_19");
