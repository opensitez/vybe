// vybe-test: csharp/csharp_collection_types/observable_collection_collection_changed_fires_on_add
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

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

var oc=new System.Collections.ObjectModel.ObservableCollection<int>();
int count=0;
oc.CollectionChanged+=(s,e)=>count++;
oc.Add(1); oc.Add(2);
__P((count).ToString());
__Check("2");
