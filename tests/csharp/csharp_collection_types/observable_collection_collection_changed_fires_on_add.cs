// vybe-test: csharp/csharp_collection_types/observable_collection_collection_changed_fires_on_add
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var oc=new System.Collections.ObjectModel.ObservableCollection<int>();
int count=0;
oc.CollectionChanged+=(s,e)=>count++;
oc.Add(1); oc.Add(2);
__Check((count).ToString(), "2");
