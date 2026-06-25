//! Collection types not elsewhere covered: SortedSet, ObservableCollection, Queue, Stack.
use super::helpers::run_csharp;

#[test]
fn sorted_set_maintains_unique_sorted_elements() {
    assert_eq!(
        run_csharp(r#"var s=new System.Collections.Generic.SortedSet<int>{3,1,4,1,5};
Console.WriteLine(s.Count);
Console.WriteLine(s.Min); Console.WriteLine(s.Max);"#),
        &["4", "1", "5"]
    );
}

#[test]
fn sorted_set_remove_eliminates_element() {
    assert_eq!(
        run_csharp(r#"var s=new System.Collections.Generic.SortedSet<int>{1,2,3};
s.Remove(2);
Console.WriteLine(s.Count);"#),
        &["2"]
    );
}

#[test]
fn observable_collection_collection_changed_fires_on_add() {
    assert_eq!(
        run_csharp(r#"var oc=new System.Collections.ObjectModel.ObservableCollection<int>();
int count=0;
oc.CollectionChanged+=(s,e)=>count++;
oc.Add(1); oc.Add(2);
Console.WriteLine(count);"#),
        &["2"]
    );
}

#[test]
fn queue_enqueue_dequeue_maintains_fifo() {
    assert_eq!(
        run_csharp(r#"var q=new System.Collections.Generic.Queue<string>();
q.Enqueue("first"); q.Enqueue("second");
Console.WriteLine(q.Dequeue());"#),
        &["first"]
    );
}

#[test]
fn stack_push_pop_maintains_lifo() {
    assert_eq!(
        run_csharp(r#"var s=new System.Collections.Generic.Stack<string>();
s.Push("a"); s.Push("b");
Console.WriteLine(s.Pop());"#),
        &["b"]
    );
}

#[test]
fn linked_list_add_after_inserts_between_nodes() {
    assert_eq!(
        run_csharp(r#"var ll=new System.Collections.Generic.LinkedList<int>();
var n1=ll.AddFirst(1);
ll.AddAfter(n1,3);
ll.AddAfter(n1,2);
Console.WriteLine(ll.First.Next.Value);"#),
        &["2"]
    );
}

#[test]
fn sorted_dictionary_first_key_is_smallest() {
    assert_eq!(
        run_csharp(r#"var sd=new System.Collections.Generic.SortedDictionary<int,string>{{3,"c"},{1,"a"},{2,"b"}};
Console.WriteLine(sd.Keys.First());"#),
        &["1"]
    );
}
