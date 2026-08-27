// vybe-test: csharp/csharp_collection_types/observable_collection_collection_changed_fires_on_add
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

using static __Harness;

var oc=new System.Collections.ObjectModel.ObservableCollection<int>();
int count=0;
oc.CollectionChanged+=(s,e)=>count++;
oc.Add(1);
oc.Add(2);
__P((count).ToString());
__Check("2");

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
