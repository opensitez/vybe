//! container/heap min-heap Push/Pop/Fix, container/list doubly-linked ops, container/ring
//! circular buffer — distinct from `test_container_list_ring.rs` (minimal smoke) and
//! `test_cover_hash_heap_io.rs` (compile-only heap Init/Push/Fix).

go_run_cases! {
    heap_push_pop_min_order => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{5, 3, 7, 1}; heap.Init(h); fmt.Println(heap.Pop(h)); fmt.Println(heap.Pop(h)) }",
        vec!["1", "3"]
    ),
    heap_push_increases_len => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{}; heap.Init(h); heap.Push(h, 4); heap.Push(h, 2); fmt.Println(h.Len()) }",
        vec!["2"]
    ),
    heap_pop_empty_after_drain => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{9, 5, 7}; heap.Init(h); heap.Pop(h); heap.Pop(h); heap.Pop(h); fmt.Println(h.Len()) }",
        vec!["0"]
    ),
    heap_init_from_unsorted => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{10, 4, 15, 20, 0, 11, 7}; heap.Init(h); fmt.Println(heap.Pop(h)); fmt.Println(heap.Pop(h)) }",
        vec!["0", "4"]
    ),
    heap_push_then_pop_returns_min => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{8, 6}; heap.Init(h); heap.Push(h, 1); fmt.Println(heap.Pop(h)) }",
        vec!["1"]
    ),
    heap_fix_after_root_decrease => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{5, 3, 7, 1}; heap.Init(h); (*h)[0] = 0; heap.Fix(h, 0); fmt.Println(heap.Pop(h)) }",
        vec!["0"]
    ),
    heap_fix_after_leaf_increase => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{1, 2, 3, 9}; heap.Init(h); (*h)[3] = 10; heap.Fix(h, 3); fmt.Println(heap.Pop(h)) }",
        vec!["1"]
    ),
    heap_remove_middle_element => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{1, 3, 5, 7, 9}; heap.Init(h); heap.Remove(h, 2); fmt.Println(h.Len()); fmt.Println(heap.Pop(h)) }",
        vec!["4", "1"]
    ),
    heap_remove_root => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{2, 4, 6, 8}; heap.Init(h); heap.Remove(h, 0); fmt.Println(heap.Pop(h)) }",
        vec!["4"]
    ),
    heap_single_element => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{42}; heap.Init(h); fmt.Println(heap.Pop(h)); fmt.Println(h.Len()) }",
        vec!["42", "0"]
    ),
    heap_pop_all_ascending => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{5, 1, 4, 2, 3}; heap.Init(h); fmt.Println(heap.Pop(h)); fmt.Println(heap.Pop(h)); fmt.Println(heap.Pop(h)) }",
        vec!["1", "2", "3"]
    ),
    heap_push_many_then_pop_min => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{}; heap.Init(h); heap.Push(h, 100); heap.Push(h, 50); heap.Push(h, 25); fmt.Println(heap.Pop(h)) }",
        vec!["25"]
    ),

    list_push_front_order => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); l.PushFront(2); l.PushFront(1); fmt.Println(l.Front().Value); fmt.Println(l.Back().Value) }",
        vec!["1", "2"]
    ),
    list_push_back_order => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); l.PushBack(1); l.PushBack(2); fmt.Println(l.Front().Value); fmt.Println(l.Back().Value) }",
        vec!["1", "2"]
    ),
    list_push_front_back_mixed => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); l.PushBack(2); l.PushFront(1); l.PushBack(3); fmt.Println(l.Front().Value); fmt.Println(l.Back().Value); fmt.Println(l.Len()) }",
        vec!["1", "3", "3"]
    ),
    list_len_after_operations => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); l.PushBack(1); l.PushBack(2); l.PushFront(0); fmt.Println(l.Len()) }",
        vec!["3"]
    ),
    list_remove_single => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); e := l.PushBack(99); l.Remove(e); fmt.Println(l.Len()); fmt.Println(l.Front() == nil) }",
        vec!["0", "true"]
    ),
    list_remove_front => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); e := l.PushFront(1); l.PushBack(2); l.Remove(e); fmt.Println(l.Front().Value) }",
        vec!["2"]
    ),
    list_remove_back => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); l.PushFront(1); e := l.PushBack(2); l.Remove(e); fmt.Println(l.Back().Value) }",
        vec!["1"]
    ),
    list_move_before_reorders => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); a := l.PushBack(1); b := l.PushBack(2); c := l.PushBack(3); l.MoveBefore(c, a); fmt.Println(l.Front().Value); fmt.Println(l.Back().Value) }",
        vec!["3", "2"]
    ),
    list_move_after_reorders => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); a := l.PushBack(1); b := l.PushBack(2); c := l.PushBack(3); l.MoveAfter(a, c); fmt.Println(l.Back().Value) }",
        vec!["1"]
    ),
    list_move_to_front => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); a := l.PushBack(1); l.PushBack(2); l.MoveBefore(a, l.Front()); fmt.Println(l.Front().Value) }",
        vec!["1"]
    ),
    list_move_to_back => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); l.PushBack(1); b := l.PushBack(2); l.MoveAfter(b, l.Back()); fmt.Println(l.Back().Value) }",
        vec!["2"]
    ),
    list_insert_before => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); e := l.PushBack(2); l.InsertBefore(1, e); fmt.Println(l.Front().Value); fmt.Println(l.Back().Value) }",
        vec!["1", "2"]
    ),
    list_insert_after => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); e := l.PushBack(1); l.InsertAfter(2, e); fmt.Println(l.Front().Value); fmt.Println(l.Back().Value) }",
        vec!["1", "2"]
    ),
    list_empty_new => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); fmt.Println(l.Len()); fmt.Println(l.Front() == nil) }",
        vec!["0", "true"]
    ),
    list_remove_all => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); e1 := l.PushBack(1); e2 := l.PushBack(2); l.Remove(e1); l.Remove(e2); fmt.Println(l.Len()) }",
        vec!["0"]
    ),
    list_string_values => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); l.PushBack(\"go\"); l.PushBack(\"vybe\"); fmt.Println(l.Front().Value); fmt.Println(l.Back().Value) }",
        vec!["go", "vybe"]
    ),
    list_move_middle_element => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); a := l.PushBack(1); b := l.PushBack(2); c := l.PushBack(3); l.MoveBefore(b, c); var vals []int; for e := l.Front(); e != nil; e = e.Next() { vals = append(vals, e.Value.(int)) }; fmt.Println(vals[0]); fmt.Println(vals[2]) }",
        vec!["1", "2"]
    ),

    ring_new_len => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(5); fmt.Println(r.Len()) }",
        vec!["5"]
    ),
    ring_len_one => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(1); fmt.Println(r.Len()) }",
        vec!["1"]
    ),
    ring_next_cycles => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(3); r.Value = 10; r = r.Next(); r.Value = 20; r = r.Next(); r.Value = 30; r = r.Next(); fmt.Println(r.Value) }",
        vec!["10"]
    ),
    ring_prev_cycles => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(3); r.Value = 1; r.Next().Value = 2; r.Next().Next().Value = 3; fmt.Println(r.Prev().Value) }",
        vec!["3"]
    ),
    ring_do_sums_values => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(4); sum := 0; for i := 0; i < 4; i++ { r.Value = i + 1; r = r.Next() }; r.Do(func(v interface{}) { sum += v.(int) }); fmt.Println(sum) }",
        vec!["10"]
    ),
    ring_do_empty_len_zero => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(0); count := 0; r.Do(func(v interface{}) { count++ }); fmt.Println(count) }",
        vec!["0"]
    ),
    ring_do_single_element => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(1); r.Value = 42; count := 0; r.Do(func(v interface{}) { count++; fmt.Println(v) }); fmt.Println(count) }",
        vec!["42", "1"]
    ),
    ring_link_combines => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { a := ring.New(2); b := ring.New(2); a.Value = 1; a.Next().Value = 2; b.Value = 3; b.Next().Value = 4; a.Link(b); fmt.Println(a.Len()) }",
        vec!["4"]
    ),
    ring_unlink_splits => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(4); split := r.Unlink(2); fmt.Println(r.Len()); fmt.Println(split.Len()) }",
        vec!["2", "2"]
    ),
    ring_move_sets_value => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(3); r.Value = 0; r.Move(2); fmt.Println(r.Value) }",
        vec!["2"]
    ),
    ring_next_after_link => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { a := ring.New(1); b := ring.New(1); a.Value = 10; b.Value = 20; a.Link(b); fmt.Println(a.Next().Value) }",
        vec!["20"]
    ),
    ring_do_string_values => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(2); r.Value = \"a\"; r.Next().Value = \"b\"; first := \"\"; r.Do(func(v interface{}) { if first == \"\" { first = v.(string) } }); fmt.Println(first) }",
        vec!["a"]
    ),
    heap_two_elements_pop_both => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{2, 1}; heap.Init(h); fmt.Println(heap.Pop(h)); fmt.Println(heap.Pop(h)) }",
        vec!["1", "2"]
    ),
    heap_fix_after_push => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{3, 1, 2}; heap.Init(h); heap.Push(h, 0); fmt.Println(heap.Pop(h)) }",
        vec!["0"]
    ),
    list_back_nil_when_empty => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); fmt.Println(l.Back() == nil) }",
        vec!["true"]
    ),
    list_iterate_forward => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); l.PushBack(1); l.PushBack(2); l.PushBack(3); sum := 0; for e := l.Front(); e != nil; e = e.Next() { sum += e.Value.(int) }; fmt.Println(sum) }",
        vec!["6"]
    ),
    list_iterate_backward => (
        "package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); l.PushBack(1); l.PushBack(2); l.PushBack(3); sum := 0; for e := l.Back(); e != nil; e = e.Prev() { sum += e.Value.(int) }; fmt.Println(sum) }",
        vec!["6"]
    ),
    ring_do_visits_each_once => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(3); count := 0; r.Do(func(v interface{}) { count++ }); fmt.Println(count) }",
        vec!["3"]
    ),
    ring_next_twice_returns_same => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(1); r.Value = 7; fmt.Println(r.Next().Value) }",
        vec!["7"]
    ),
    ring_prev_on_single => (
        "package main; import \"fmt\"; import \"container/ring\"; func main() { r := ring.New(1); r.Value = 9; fmt.Println(r.Prev().Value) }",
        vec!["9"]
    ),
    heap_remove_then_pop_min => (
        "package main; import \"fmt\"; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{1, 5, 3, 7, 2}; heap.Init(h); heap.Remove(h, 3); fmt.Println(heap.Pop(h)); fmt.Println(heap.Pop(h)) }",
        vec!["1", "2"]
    ),
}

go_compile_cases! {
    heap_empty_init => "package main; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{}; heap.Init(h) }",
    heap_fix_leaf_index => "package main; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{1, 2, 3}; heap.Init(h); heap.Fix(h, 1) }",
    list_push_back_list => "package main; import \"container/list\"; func main() { a := list.New(); b := list.New(); b.PushBack(1); a.PushBackList(b) }",
    list_push_front_list => "package main; import \"container/list\"; func main() { a := list.New(); b := list.New(); b.PushFront(1); a.PushFrontList(b) }",
    list_init_clears => "package main; import \"container/list\"; func main() { l := list.New(); l.PushBack(1); l.Init() }",
    ring_new_zero => "package main; import \"container/ring\"; func main() { r := ring.New(0); _ = r.Next() }",
    ring_link_self => "package main; import \"container/ring\"; func main() { r := ring.New(3); r.Link(r) }",
    list_back_nil_when_empty_compile => "package main; import \"container/list\"; func main() { l := list.New(); _ = l.Back() }",
    heap_remove_last => "package main; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{1, 2, 3}; heap.Init(h); heap.Remove(h, 2) }",
    ring_prev_on_two => "package main; import \"container/ring\"; func main() { r := ring.New(2); _ = r.Prev() }",
    list_move_after_self => "package main; import \"container/list\"; func main() { l := list.New(); e := l.PushBack(1); l.MoveAfter(e, e) }",
    heap_push_duplicate_values => "package main; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{}; heap.Init(h); heap.Push(h, 5); heap.Push(h, 5) }",
    ring_do_mutate => "package main; import \"container/ring\"; func main() { r := ring.New(2); r.Do(func(v interface{}) {}) }",
    list_insert_before_root => "package main; import \"container/list\"; func main() { l := list.New(); e := l.PushBack(1); l.InsertBefore(0, e) }",
    heap_negative_values => "package main; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{-1, -5, 0}; heap.Init(h); _ = heap.Pop(h) }",
    ring_unlink_zero => "package main; import \"container/ring\"; func main() { r := ring.New(3); r.Unlink(0) }",
    list_remove_middle => "package main; import \"container/list\"; func main() { l := list.New(); l.PushBack(1); e := l.PushBack(2); l.PushBack(3); l.Remove(e) }",
    heap_fix_root_only => "package main; import \"container/heap\"; type IH []int; func (h IH) Len() int { return len(h) }; func (h IH) Less(i, j int) bool { return h[i] < h[j] }; func (h IH) Swap(i, j int) { h[i], h[j] = h[j], h[i] }; func (h *IH) Push(x interface{}) { *h = append(*h, x.(int)) }; func (h *IH) Pop() interface{} { o := *h; n := len(o); x := o[n-1]; *h = o[:n-1]; return x }; func main() { h := &IH{3, 1, 2}; heap.Init(h); heap.Fix(h, 0) }",
    ring_move_negative => "package main; import \"container/ring\"; func main() { r := ring.New(4); r.Move(-1) }",
    list_push_back_nil_list => "package main; import \"container/list\"; func main() { a := list.New(); a.PushBackList(list.New()) }",
}
