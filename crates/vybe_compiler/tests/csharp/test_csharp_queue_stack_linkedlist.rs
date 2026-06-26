//! Queue<T>, Stack<T>, and LinkedList<T> operational semantics beyond basic FIFO/LIFO.

csharp_cases! {
    queue_enqueue_increments_count => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(1); q.Enqueue(2); Console.WriteLine(q.Count);"#,
        ["2"]
    };

    queue_dequeue_returns_oldest_element => {
        r#"using System.Collections.Generic; var q = new Queue<string>(); q.Enqueue("first"); q.Enqueue("second"); Console.WriteLine(q.Dequeue());"#,
        ["first"]
    };

    queue_dequeue_twice_drains_in_fifo_order => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(10); q.Enqueue(20); q.Enqueue(30); Console.WriteLine(q.Dequeue()); Console.WriteLine(q.Dequeue());"#,
        ["10", "20"]
    };

    queue_peek_reads_head_without_removing => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(5); q.Enqueue(6); Console.WriteLine(q.Peek()); Console.WriteLine(q.Count);"#,
        ["5", "2"]
    };

    queue_peek_after_dequeue_shows_next_head => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(1); q.Enqueue(2); q.Dequeue(); Console.WriteLine(q.Peek());"#,
        ["2"]
    };

    queue_contains_finds_enqueued_value => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(7); q.Enqueue(8); Console.WriteLine(q.Contains(8));"#,
        ["True"]
    };

    queue_contains_reports_absent_value => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(1); Console.WriteLine(q.Contains(99));"#,
        ["False"]
    };

    queue_clear_empties_all_elements => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(1); q.Enqueue(2); q.Clear(); Console.WriteLine(q.Count);"#,
        ["0"]
    };

    queue_enqueue_after_clear_starts_fresh => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(1); q.Clear(); q.Enqueue(9); Console.WriteLine(q.Dequeue());"#,
        ["9"]
    };

    queue_single_element_peek_equals_dequeue => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(42); Console.WriteLine(q.Peek()); Console.WriteLine(q.Dequeue());"#,
        ["42", "42"]
    };

    queue_to_array_preserves_fifo_order => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(3); q.Enqueue(1); q.Enqueue(2); var arr = q.ToArray(); Console.WriteLine(arr[0]); Console.WriteLine(arr[2]);"#,
        ["3", "2"]
    };

    queue_dequeue_reduces_count_by_one => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(1); q.Enqueue(2); q.Dequeue(); Console.WriteLine(q.Count);"#,
        ["1"]
    };

    queue_string_elements_maintain_insertion_order => {
        r#"using System.Collections.Generic; var q = new Queue<string>(); q.Enqueue("a"); q.Enqueue("b"); Console.WriteLine(q.Dequeue()); Console.WriteLine(q.Dequeue());"#,
        ["a", "b"]
    };

    stack_push_increments_count => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Push(2); Console.WriteLine(s.Count);"#,
        ["2"]
    };

    stack_pop_returns_most_recently_pushed => {
        r#"using System.Collections.Generic; var s = new Stack<string>(); s.Push("bottom"); s.Push("top"); Console.WriteLine(s.Pop());"#,
        ["top"]
    };

    stack_pop_twice_drains_in_lifo_order => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Push(2); s.Push(3); Console.WriteLine(s.Pop()); Console.WriteLine(s.Pop());"#,
        ["3", "2"]
    };

    stack_peek_reads_top_without_removing => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(4); s.Push(5); Console.WriteLine(s.Peek()); Console.WriteLine(s.Count);"#,
        ["5", "2"]
    };

    stack_peek_after_pop_shows_new_top => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Push(2); s.Pop(); Console.WriteLine(s.Peek());"#,
        ["1"]
    };

    stack_contains_finds_pushed_value => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(11); s.Push(22); Console.WriteLine(s.Contains(22));"#,
        ["True"]
    };

    stack_contains_reports_absent_value => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(5); Console.WriteLine(s.Contains(0));"#,
        ["False"]
    };

    stack_clear_empties_all_elements => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Push(2); s.Clear(); Console.WriteLine(s.Count);"#,
        ["0"]
    };

    stack_push_after_clear_restarts_sequence => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Clear(); s.Push(8); Console.WriteLine(s.Pop());"#,
        ["8"]
    };

    stack_single_element_peek_equals_pop => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(99); Console.WriteLine(s.Peek()); Console.WriteLine(s.Pop());"#,
        ["99", "99"]
    };

    stack_to_array_preserves_lifo_top_at_end => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Push(2); s.Push(3); var arr = s.ToArray(); Console.WriteLine(arr[0]); Console.WriteLine(arr[2]);"#,
        ["3", "1"]
    };

    stack_pop_reduces_count_by_one => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Push(2); s.Pop(); Console.WriteLine(s.Count);"#,
        ["1"]
    };

    linkedlist_add_first_inserts_at_head => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(2); ll.AddFirst(1); Console.WriteLine(ll.First.Value);"#,
        ["1"]
    };

    linkedlist_add_last_appends_at_tail => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddFirst(1); ll.AddLast(3); Console.WriteLine(ll.Last.Value);"#,
        ["3"]
    };

    linkedlist_add_after_inserts_between_nodes => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); var n = ll.AddFirst(1); ll.AddAfter(n, 3); ll.AddAfter(n, 2); Console.WriteLine(n.Next.Value);"#,
        ["2"]
    };

    linkedlist_add_before_inserts_predecessor => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); var tail = ll.AddLast(3); ll.AddBefore(tail, 2); ll.AddBefore(tail, 1); Console.WriteLine(ll.First.Value);"#,
        ["1"]
    };

    linkedlist_remove_node_by_reference => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); var mid = ll.AddLast(2); ll.AddFirst(1); ll.AddLast(3); ll.Remove(mid); Console.WriteLine(ll.Count);"#,
        ["2"]
    };

    linkedlist_remove_first_drops_head => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddLast(2); ll.RemoveFirst(); Console.WriteLine(ll.First.Value);"#,
        ["2"]
    };

    linkedlist_remove_last_drops_tail => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddLast(2); ll.RemoveLast(); Console.WriteLine(ll.Last.Value);"#,
        ["1"]
    };

    linkedlist_find_returns_matching_node => {
        r#"using System.Collections.Generic; var ll = new LinkedList<string>(); ll.AddLast("a"); ll.AddLast("target"); var node = ll.Find("target"); Console.WriteLine(node.Value);"#,
        ["target"]
    };

    linkedlist_find_returns_null_for_missing => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); Console.WriteLine(ll.Find(9) == null);"#,
        ["True"]
    };

    linkedlist_contains_detects_present_value => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(5); ll.AddLast(6); Console.WriteLine(ll.Contains(6));"#,
        ["True"]
    };

    linkedlist_contains_false_for_absent_value => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); Console.WriteLine(ll.Contains(2));"#,
        ["False"]
    };

    linkedlist_count_tracks_additions => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddLast(2); ll.AddLast(3); Console.WriteLine(ll.Count);"#,
        ["3"]
    };

    linkedlist_clear_removes_all_nodes => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddLast(2); ll.Clear(); Console.WriteLine(ll.Count);"#,
        ["0"]
    };

    linkedlist_foreach_walks_forward_order => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddLast(2); ll.AddLast(3); foreach (var x in ll) Console.WriteLine(x);"#,
        ["1", "2", "3"]
    };

    linkedlist_first_next_links_to_second_element => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(10); ll.AddLast(20); Console.WriteLine(ll.First.Next.Value);"#,
        ["20"]
    };

    linkedlist_last_previous_links_to_penultimate => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(10); ll.AddLast(20); ll.AddLast(30); Console.WriteLine(ll.Last.Previous.Value);"#,
        ["20"]
    };

    linkedlist_add_first_on_empty_becomes_sole_node => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddFirst(7); Console.WriteLine(ll.First.Value); Console.WriteLine(ll.Last.Value);"#,
        ["7", "7"]
    };

    linkedlist_remove_value_by_payload => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddLast(2); ll.AddLast(3); ll.Remove(2); Console.WriteLine(ll.Contains(2));"#,
        ["False"]
    };

    queue_three_step_enqueue_dequeue_cycle => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(1); q.Enqueue(2); q.Enqueue(3); q.Dequeue(); Console.WriteLine(q.Peek()); Console.WriteLine(q.Count);"#,
        ["2", "2"]
    };

    stack_three_push_pop_leaves_one => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Push(2); s.Push(3); s.Pop(); s.Pop(); Console.WriteLine(s.Pop());"#,
        ["1"]
    };

    linkedlist_add_last_on_empty_sets_head_and_tail => {
        r#"using System.Collections.Generic; var ll = new LinkedList<string>(); ll.AddLast("solo"); Console.WriteLine(ll.First.Value); Console.WriteLine(ll.Last.Value);"#,
        ["solo", "solo"]
    };

    linkedlist_remove_first_on_singleton_empties_list => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(9); ll.RemoveFirst(); Console.WriteLine(ll.Count);"#,
        ["0"]
    };

    queue_empty_after_full_drain => {
        r#"using System.Collections.Generic; var q = new Queue<int>(); q.Enqueue(1); q.Dequeue(); Console.WriteLine(q.Count);"#,
        ["0"]
    };

    stack_empty_after_full_drain => {
        r#"using System.Collections.Generic; var s = new Stack<int>(); s.Push(1); s.Pop(); Console.WriteLine(s.Count);"#,
        ["0"]
    };

    linkedlist_add_after_last_node_extends_tail => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddAfter(ll.Last, 2); Console.WriteLine(ll.Last.Value);"#,
        ["2"]
    };

    linkedlist_add_before_first_node_extends_head => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(2); ll.AddBefore(ll.First, 1); Console.WriteLine(ll.First.Value);"#,
        ["1"]
    };

    queue_bool_elements_roundtrip => {
        r#"using System.Collections.Generic; var q = new Queue<bool>(); q.Enqueue(true); q.Enqueue(false); Console.WriteLine(q.Dequeue()); Console.WriteLine(q.Dequeue());"#,
        ["True", "False"]
    };

    stack_string_elements_lifo_order => {
        r#"using System.Collections.Generic; var s = new Stack<string>(); s.Push("x"); s.Push("y"); Console.WriteLine(s.Pop()); Console.WriteLine(s.Pop());"#,
        ["y", "x"]
    };

    linkedlist_find_returns_first_of_duplicates => {
        r#"using System.Collections.Generic; var ll = new LinkedList<int>(); ll.AddLast(1); ll.AddLast(2); ll.AddLast(2); var node = ll.Find(2); Console.WriteLine(node.Value); Console.WriteLine(node.Next.Value);"#,
        ["2", "2"]
    };
}
