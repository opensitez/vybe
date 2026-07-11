//! context cancel/deadline: WithCancel, WithDeadline, WithTimeout, WithValue chains, Done, Err sentinels.

go_run_cases! {
    context_canceled_not_equal_deadline_exceeded => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.Canceled != context.DeadlineExceeded) }",
        vec!["true"]
    ),
    context_background_err_nil => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.Background().Err() == nil) }",
        vec!["true"]
    ),
    context_todo_err_nil => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.TODO().Err() == nil) }",
        vec!["true"]
    ),
    context_background_done_nil => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.Background().Done() == nil) }",
        vec!["true"]
    ),
    context_todo_done_nil => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.TODO().Done() == nil) }",
        vec!["true"]
    ),
    context_with_cancel_err_after_cancel_func => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); cancel(); fmt.Println(ctx.Err() == context.Canceled) }",
        vec!["true"]
    ),
    context_with_cancel_done_closed_after_cancel => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); cancel(); select { case <-ctx.Done(): fmt.Println(\"done\"); default: fmt.Println(\"pending\") } }",
        vec!["done"]
    ),
    context_with_cancel_done_open_before_cancel => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx, _ := context.WithCancel(context.Background()); select { case <-ctx.Done(): fmt.Println(\"done\"); default: fmt.Println(\"pending\") } }",
        vec!["pending"]
    ),
    context_with_value_three_level_chain => (
        "package main; import \"fmt\"; import \"context\"; func main() { c1 := context.WithValue(context.Background(), \"a\", 1); c2 := context.WithValue(c1, \"b\", 2); c3 := context.WithValue(c2, \"c\", 3); fmt.Println(c3.Value(\"a\").(int)); fmt.Println(c3.Value(\"b\").(int)); fmt.Println(c3.Value(\"c\").(int)) }",
        vec!["1", "2", "3"]
    ),
    context_with_value_parent_visible_from_child => (
        "package main; import \"fmt\"; import \"context\"; func main() { parent := context.WithValue(context.Background(), \"trace\", \"root\"); child, _ := context.WithCancel(parent); fmt.Println(child.Value(\"trace\").(string)) }",
        vec!["root"]
    ),
    context_with_value_child_shadows_parent_key => (
        "package main; import \"fmt\"; import \"context\"; func main() { parent := context.WithValue(context.Background(), \"k\", \"old\"); child := context.WithValue(parent, \"k\", \"new\"); fmt.Println(child.Value(\"k\").(string)) }",
        vec!["new"]
    ),
    context_with_deadline_err_nil_before_expiry => (
        "package main; import \"fmt\"; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithDeadline(context.Background(), time.Now().Add(time.Hour)); defer cancel(); fmt.Println(ctx.Err() == nil) }",
        vec!["true"]
    ),
    context_with_deadline_has_ok_true => (
        "package main; import \"fmt\"; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithDeadline(context.Background(), time.Now().Add(time.Hour)); defer cancel(); _, ok := ctx.Deadline(); fmt.Println(ok) }",
        vec!["true"]
    ),
    context_with_deadline_manual_cancel_is_canceled => (
        "package main; import \"fmt\"; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithDeadline(context.Background(), time.Now().Add(time.Hour)); cancel(); fmt.Println(ctx.Err() == context.Canceled) }",
        vec!["true"]
    ),
    context_with_timeout_deadline_exceeded_after_sleep => (
        "package main; import \"fmt\"; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Millisecond); defer cancel(); time.Sleep(2 * time.Millisecond); fmt.Println(ctx.Err() == context.DeadlineExceeded) }",
        vec!["true"]
    ),
    context_with_timeout_err_nil_immediately => (
        "package main; import \"fmt\"; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Minute); defer cancel(); fmt.Println(ctx.Err() == nil) }",
        vec!["true"]
    ),
    context_canceled_error_string => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.Canceled.Error()) }",
        vec!["context canceled"]
    ),
    context_deadline_exceeded_error_string => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.DeadlineExceeded.Error()) }",
        vec!["deadline exceeded"]
    ),
    context_with_cancel_parent_canceled_propagates => (
        "package main; import \"fmt\"; import \"context\"; func main() { parent, pcancel := context.WithCancel(context.Background()); child, _ := context.WithCancel(parent); pcancel(); fmt.Println(child.Err() == context.Canceled) }",
        vec!["true"]
    ),
    context_with_cancel_child_only_parent_still_active => (
        "package main; import \"fmt\"; import \"context\"; func main() { parent, _ := context.WithCancel(context.Background()); child, ccancel := context.WithCancel(parent); ccancel(); fmt.Println(parent.Err() == nil) }",
        vec!["true"]
    ),
    context_with_value_missing_key_nil => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx := context.WithValue(context.Background(), \"x\", 1); fmt.Println(ctx.Value(\"y\") == nil) }",
        vec!["true"]
    ),
    context_with_timeout_zero_duration_expires => (
        "package main; import \"fmt\"; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithTimeout(context.Background(), 0); defer cancel(); fmt.Println(ctx.Err() == context.DeadlineExceeded) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    context_with_cancel_returns_cancel_func => "package main; import \"context\"; func main() { _, cancel := context.WithCancel(context.Background()); cancel() }",
    context_with_cancel_defer_cancel => "package main; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); defer cancel(); _ = ctx }",
    context_with_cancel_on_timeout_parent => "package main; import \"context\"; import \"time\"; func main() { parent, pcancel := context.WithTimeout(context.Background(), time.Minute); defer pcancel(); _, ccancel := context.WithCancel(parent); defer ccancel() }",
    context_with_cancel_stored_in_struct => "package main; import \"context\"; type holder struct { ctx context.Context; cancel context.CancelFunc }; func main() { h := holder{}; h.ctx, h.cancel = context.WithCancel(context.Background()); defer h.cancel() }",
    context_with_deadline_absolute_time => "package main; import \"context\"; import \"time\"; func main() { _, cancel := context.WithDeadline(context.Background(), time.Now().Add(10*time.Second)); defer cancel() }",
    context_with_deadline_past_time => "package main; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithDeadline(context.Background(), time.Now().Add(-time.Second)); defer cancel(); _ = ctx.Err() }",
    context_with_deadline_done_channel => "package main; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithDeadline(context.Background(), time.Now().Add(time.Millisecond)); defer cancel(); <-ctx.Done() }",
    context_with_timeout_one_second => "package main; import \"context\"; import \"time\"; func main() { _, cancel := context.WithTimeout(context.Background(), time.Second); defer cancel() }",
    context_with_timeout_nested_child => "package main; import \"context\"; import \"time\"; func main() { parent, pcancel := context.WithTimeout(context.Background(), time.Minute); defer pcancel(); _, ccancel := context.WithTimeout(parent, time.Second); defer ccancel() }",
    context_with_value_int_key_type => "package main; import \"context\"; type ctxKey int; func main() { const k ctxKey = 0; _ = context.WithValue(context.Background(), k, \"v\") }",
    context_with_value_struct_key_empty => "package main; import \"context\"; type key struct{}; func main() { _ = context.WithValue(context.Background(), key{}, true) }",
    context_with_value_chain_four_levels => "package main; import \"context\"; func main() { c := context.Background(); c = context.WithValue(c, \"l1\", 1); c = context.WithValue(c, \"l2\", 2); c = context.WithValue(c, \"l3\", 3); c = context.WithValue(c, \"l4\", 4); _ = c.Value(\"l2\") }",
    context_with_value_on_todo_parent => "package main; import \"context\"; func main() { _ = context.WithValue(context.TODO(), \"k\", \"v\") }",
    context_with_value_on_canceled_ctx => "package main; import \"context\"; func main() { parent, cancel := context.WithCancel(context.Background()); cancel(); _ = context.WithValue(parent, \"k\", 1) }",
    context_done_channel_receive_compile => "package main; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); go func() { cancel() }(); _ = <-ctx.Done() }",
    context_done_select_two_cases => "package main; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); defer cancel(); ch := make(chan int); select { case <-ctx.Done(): case <-ch: } }",
    context_err_compare_canceled_sentinel => "package main; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); cancel(); _ = ctx.Err() == context.Canceled }",
    context_err_compare_deadline_sentinel => "package main; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithTimeout(context.Background(), 0); defer cancel(); _ = ctx.Err() == context.DeadlineExceeded }",
    context_background_value_always_nil => "package main; import \"context\"; func main() { _ = context.Background().Value(\"any\") == nil }",
    context_todo_value_always_nil => "package main; import \"context\"; func main() { _ = context.TODO().Value(\"any\") == nil }",
    context_without_cancel_shields_child => "package main; import \"context\"; func main() { parent, pcancel := context.WithCancel(context.Background()); child := context.WithoutCancel(parent); pcancel(); _ = child.Err() == nil }",
    context_cause_on_cancel_go121 => "package main; import \"context\"; import \"errors\"; func main() { ctx, cancel := context.WithCancelCause(context.Background()); cancel(errors.New(\"stop\")); _ = context.Cause(ctx) }",
    context_after_func_on_cancel => "package main; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); context.AfterFunc(ctx, func() {}); cancel() }",
    context_with_deadline_child_inherits => "package main; import \"context\"; import \"time\"; func main() { parent, pcancel := context.WithDeadline(context.Background(), time.Now().Add(time.Hour)); defer pcancel(); child, ccancel := context.WithCancel(parent); defer ccancel(); _, ok := child.Deadline(); _ = ok }",
    context_with_timeout_passed_to_function => "package main; import \"context\"; import \"time\"; func work(ctx context.Context) { _ = ctx.Err() }; func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Second); defer cancel(); work(ctx) }",
    context_with_cancel_grandchild_chain => "package main; import \"context\"; func main() { a, ca := context.WithCancel(context.Background()); b, _ := context.WithCancel(a); _, _ = context.WithCancel(b); ca() }",
    context_with_value_readonly_after_cancel => "package main; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.WithValue(context.Background(), \"k\", 9)); cancel(); _ = ctx.Value(\"k\") }",
    context_deadline_returns_time_value => "package main; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithDeadline(context.Background(), time.Now().Add(time.Second)); defer cancel(); t, ok := ctx.Deadline(); _ = t; _ = ok }",
    context_with_timeout_cancel_before_expiry => "package main; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Hour); cancel(); _ = ctx.Err() == context.Canceled }",
    context_nested_with_value_and_cancel => "package main; import \"context\"; func main() { base := context.WithValue(context.Background(), \"id\", \"1\"); mid, mcancel := context.WithCancel(base); defer mcancel(); leaf := context.WithValue(mid, \"step\", 2); _ = leaf.Value(\"id\"); _ = leaf.Value(\"step\") }",
    context_select_default_when_not_canceled => "package main; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); defer cancel(); select { case <-ctx.Done(): default: _ = 1 } }",
    context_with_cancel_from_background => "package main; import \"context\"; func main() { _, cancel := context.WithCancel(context.Background()); defer cancel() }",
    context_with_cancel_from_todo => "package main; import \"context\"; func main() { _, cancel := context.WithCancel(context.TODO()); defer cancel() }",
    context_with_deadline_from_todo => "package main; import \"context\"; import \"time\"; func main() { _, cancel := context.WithDeadline(context.TODO(), time.Now().Add(time.Second)); defer cancel() }",
    context_with_timeout_from_todo => "package main; import \"context\"; import \"time\"; func main() { _, cancel := context.WithTimeout(context.TODO(), time.Second); defer cancel() }",
    context_done_not_nil_after_cancel => "package main; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); cancel(); _ = ctx.Done() != nil }",
    context_err_nil_on_fresh_with_cancel => "package main; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); defer cancel(); _ = ctx.Err() == nil }",
    context_with_value_bool_payload => "package main; import \"context\"; func main() { _ = context.WithValue(context.Background(), \"flag\", true).Value(\"flag\") }",
}
