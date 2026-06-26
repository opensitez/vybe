//! container/list and container/ring compile coverage.

use crate::helpers::*;

go_run_cases! {
    list_push_back_len => ("package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); l.PushBack(1); l.PushBack(2); fmt.Println(l.Len()) }", vec!["2"]),
    list_front_back => ("package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); l.PushFront(\"a\"); l.PushBack(\"b\"); fmt.Println(l.Front().Value, l.Back().Value) }", vec!["a b"]),
    list_remove => ("package main; import \"fmt\"; import \"container/list\"; func main() { l := list.New(); e := l.PushBack(9); l.Remove(e); fmt.Println(l.Len()) }", vec!["0"]),
}

go_compile_cases! {
    list_push_front => "package main; import \"container/list\"; func main() { l := list.New(); l.PushFront(1) }",
    list_insert_before => "package main; import \"container/list\"; func main() { l := list.New(); e := l.PushBack(1); l.InsertBefore(0, e) }",
    list_insert_after => "package main; import \"container/list\"; func main() { l := list.New(); e := l.PushBack(1); l.InsertAfter(2, e) }",
    list_move_before => "package main; import \"container/list\"; func main() { l := list.New(); a := l.PushBack(1); b := l.PushBack(2); l.MoveBefore(b, a) }",
    list_move_after => "package main; import \"container/list\"; func main() { l := list.New(); a := l.PushBack(1); b := l.PushBack(2); l.MoveAfter(a, b) }",
    list_push_back_list => "package main; import \"container/list\"; func main() { a := list.New(); b := list.New(); a.PushBackList(b) }",
    list_push_front_list => "package main; import \"container/list\"; func main() { a := list.New(); b := list.New(); a.PushFrontList(b) }",
    ring_new_len => "package main; import \"container/ring\"; func main() { r := ring.New(3); _ = r.Len() }",
    ring_link_unlink => "package main; import \"container/ring\"; func main() { r := ring.New(2); s := ring.New(2); r.Link(s); r.Unlink(1) }",
    ring_do => "package main; import \"container/ring\"; func main() { r := ring.New(3); r.Do(func(x interface{}) {}) }",
    ring_next_prev => "package main; import \"container/ring\"; func main() { r := ring.New(2); _ = r.Next(); _ = r.Prev() }",
}
