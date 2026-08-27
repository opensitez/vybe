// vybe-test: csharp/csharp_collections_observable_collection_events/observable_collection_case_18

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

var oc = new System.Collections.ObjectModel.ObservableCollection<string>();
bool changed = false;
oc.CollectionChanged += (s, e) => changed = true;
oc.Add("Item_18");
__P(changed.ToString());
__P(oc[0]);
__Check("True\nItem_18");
