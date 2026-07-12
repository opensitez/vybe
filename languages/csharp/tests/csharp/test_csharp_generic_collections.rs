//! Generic collection types not covered elsewhere: SortedList, PriorityQueue, ImmutableList.
use super::helpers::run_csharp;

#[test]
fn sorted_list_maintains_key_order_on_insertion() {
    assert_eq!(
        run_csharp(
            r#"var sl = new System.Collections.Generic.SortedList<int,string>();
sl.Add(3,"c"); sl.Add(1,"a"); sl.Add(2,"b");
Console.WriteLine(sl.Keys[0]);
Console.WriteLine(sl.Values[0]);"#
        ),
        &["1", "a"]
    );
}

#[test]
fn sorted_list_index_of_key_finds_insertion_position() {
    assert_eq!(
        run_csharp(
            r#"var sl = new System.Collections.Generic.SortedList<string,int>{{"a",1},{"b",2},{"c",3}};
Console.WriteLine(sl.IndexOfKey("b"));"#
        ),
        &["1"]
    );
}

#[test]
fn priority_queue_dequeue_returns_lowest_priority_first() {
    assert_eq!(
        run_csharp(
            r#"var pq = new System.Collections.Generic.PriorityQueue<string,int>();
pq.Enqueue("low", 10);
pq.Enqueue("high", 1);
pq.Enqueue("mid", 5);
Console.WriteLine(pq.Dequeue());"#
        ),
        &["high"]
    );
}

#[test]
fn priority_queue_count_decreases_after_dequeue() {
    assert_eq!(
        run_csharp(
            r#"var pq = new System.Collections.Generic.PriorityQueue<int,int>();
pq.Enqueue(1,1); pq.Enqueue(2,2);
pq.Dequeue();
Console.WriteLine(pq.Count);"#
        ),
        &["1"]
    );
}

#[test]
fn immutable_list_add_returns_new_list_without_mutating_original() {
    assert_eq!(
        run_csharp(
            r#"var original = System.Collections.Immutable.ImmutableList.Create(1,2,3);
var extended = original.Add(4);
Console.WriteLine(original.Count);
Console.WriteLine(extended.Count);"#
        ),
        &["3", "4"]
    );
}

#[test]
fn immutable_array_indexer_reads_element() {
    assert_eq!(
        run_csharp(
            r#"var arr = System.Collections.Immutable.ImmutableArray.Create(10,20,30);
Console.WriteLine(arr[1]);"#
        ),
        &["20"]
    );
}
