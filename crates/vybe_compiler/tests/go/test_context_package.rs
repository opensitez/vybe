//! context package: Background, WithCancel, WithTimeout, WithValue, Done channel.


go_run_cases! {
    background_err_is_nil => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.Background().Err() == nil) }",
        vec!["true"]
    ),
    background_value_returns_nil_for_unknown_key => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.Background().Value(\"missing\") == nil) }",
        vec!["true"]
    ),
    background_done_nil_falls_through_select_default => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx := context.Background(); select { case <-ctx.Done(): fmt.Println(\"closed\"); default: fmt.Println(\"open\") } }",
        vec!["open"]
    ),
    todo_err_is_nil => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.TODO().Err() == nil) }",
        vec!["true"]
    ),
    context_canceled_sentinel_error_string => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.Canceled.Error()) }",
        vec!["context canceled"]
    ),
    context_deadline_exceeded_sentinel_error_string => (
        "package main; import \"fmt\"; import \"context\"; func main() { fmt.Println(context.DeadlineExceeded.Error()) }",
        vec!["deadline exceeded"]
    ),
    with_value_returns_stored_string_value => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx := context.WithValue(context.Background(), \"token\", \"abc\"); fmt.Println(ctx.Value(\"token\").(string)) }",
        vec!["abc"]
    ),
    with_value_int_key_stored_and_retrieved => (
        "package main; import \"fmt\"; import \"context\"; func main() { type key int; const idKey key = 1; ctx := context.WithValue(context.Background(), idKey, 99); fmt.Println(ctx.Value(idKey).(int)) }",
        vec!["99"]
    ),
    with_value_missing_key_returns_nil => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx := context.WithValue(context.Background(), \"a\", 1); fmt.Println(ctx.Value(\"b\") == nil) }",
        vec!["true"]
    ),
    with_value_child_overrides_parent_same_key => (
        "package main; import \"fmt\"; import \"context\"; func main() { parent := context.WithValue(context.Background(), \"k\", 1); child := context.WithValue(parent, \"k\", 2); fmt.Println(child.Value(\"k\").(int)) }",
        vec!["2"]
    ),
    with_value_child_inherits_parent_other_key => (
        "package main; import \"fmt\"; import \"context\"; func main() { parent := context.WithValue(context.Background(), \"x\", 10); child := context.WithValue(parent, \"y\", 20); fmt.Println(child.Value(\"x\").(int)); fmt.Println(child.Value(\"y\").(int)) }",
        vec!["10", "20"]
    ),
    with_cancel_err_nil_before_cancel => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx, _ := context.WithCancel(context.Background()); fmt.Println(ctx.Err() == nil) }",
        vec!["true"]
    ),
    with_cancel_err_equals_canceled_after_cancel => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); cancel(); fmt.Println(ctx.Err() == context.Canceled) }",
        vec!["true"]
    ),
    with_cancel_done_select_default_before_cancel => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx, _ := context.WithCancel(context.Background()); select { case <-ctx.Done(): fmt.Println(\"closed\"); default: fmt.Println(\"open\") } }",
        vec!["open"]
    ),
    with_cancel_done_select_receives_after_cancel => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); cancel(); select { case <-ctx.Done(): fmt.Println(\"closed\"); default: fmt.Println(\"open\") } }",
        vec!["closed"]
    ),
    with_cancel_child_canceled_when_parent_canceled => (
        "package main; import \"fmt\"; import \"context\"; func main() { parent, parentCancel := context.WithCancel(context.Background()); child, _ := context.WithCancel(parent); parentCancel(); fmt.Println(child.Err() == context.Canceled) }",
        vec!["true"]
    ),
    with_cancel_local_cancel_leaves_parent_active => (
        "package main; import \"fmt\"; import \"context\"; func main() { parent, _ := context.WithCancel(context.Background()); child, childCancel := context.WithCancel(parent); childCancel(); fmt.Println(parent.Err() == nil); fmt.Println(child.Err() == context.Canceled) }",
        vec!["true", "true"]
    ),
    with_cancel_double_cancel_still_canceled => (
        "package main; import \"fmt\"; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); cancel(); cancel(); fmt.Println(ctx.Err() == context.Canceled) }",
        vec!["true"]
    ),
    with_cancel_retains_parent_with_value => (
        "package main; import \"fmt\"; import \"context\"; func main() { parent := context.WithValue(context.Background(), \"trace\", \"t1\"); child, _ := context.WithCancel(parent); fmt.Println(child.Value(\"trace\").(string)) }",
        vec!["t1"]
    ),
    errors_is_matches_context_canceled => (
        "package main; import \"fmt\"; import \"context\"; import \"errors\"; func main() { ctx, cancel := context.WithCancel(context.Background()); cancel(); fmt.Println(errors.Is(ctx.Err(), context.Canceled)) }",
        vec!["true"]
    ),
    with_timeout_manual_cancel_err_is_canceled => (
        "package main; import \"fmt\"; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Minute); cancel(); fmt.Println(ctx.Err() == context.Canceled) }",
        vec!["true"]
    ),
    with_timeout_expires_to_deadline_exceeded => (
        "package main; import \"fmt\"; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Millisecond); defer cancel(); time.Sleep(2 * time.Millisecond); fmt.Println(ctx.Err() == context.DeadlineExceeded) }",
        vec!["true"]
    ),
    without_cancel_shields_child_from_parent_cancel => (
        "package main; import \"fmt\"; import \"context\"; func main() { parent, parentCancel := context.WithCancel(context.Background()); child := context.WithoutCancel(parent); parentCancel(); fmt.Println(child.Err() == nil) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    with_deadline_absolute_time => "package main; import \"context\"; import \"time\"; func main() { _, cancel := context.WithDeadline(context.Background(), time.Now().Add(time.Second)); defer cancel() }",
    with_timeout_goroutine_select_on_done => "package main; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Second); defer cancel(); go func() { select { case <-ctx.Done(): } }() }",
    context_passed_to_helper_function => "package main; import \"context\"; func work(ctx context.Context) error { return ctx.Err() }; func main() { _ = work(context.Background()) }",
    nested_with_cancel_grandchild_chain => "package main; import \"context\"; func main() { a, ca := context.WithCancel(context.Background()); b, _ := context.WithCancel(a); _, _ = context.WithCancel(b); ca() }",
    defer_cancel_on_with_timeout => "package main; import \"context\"; import \"time\"; func main() { ctx, cancel := context.WithTimeout(context.Background(), time.Hour); defer cancel(); _ = ctx }",
    with_value_custom_struct_key => "package main; import \"context\"; type ctxKey struct{}; func main() { _ = context.WithValue(context.Background(), ctxKey{}, 1) }",
    select_ctx_done_or_work_channel => "package main; import \"context\"; func main() { ctx, cancel := context.WithCancel(context.Background()); defer cancel(); work := make(chan int, 1); select { case <-ctx.Done(): case work <- 1: } }",
}
