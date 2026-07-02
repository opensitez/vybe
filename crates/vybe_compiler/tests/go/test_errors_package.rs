//! errors package: New, fmt.Errorf wrapping, Is, As, Unwrap, Join.


go_run_cases! {
    errors_new_error_string => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.New(\"file not found\"); fmt.Println(err.Error()) }",
        vec!["file not found"]
    ),
    errors_new_distinct_instances => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.New(\"x\") == errors.New(\"x\")) }",
        vec!["false"]
    ),
    errors_new_value_not_nil => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.New(\"fail\"); fmt.Println(err != nil) }",
        vec!["true"]
    ),
    errors_sentinel_same_variable_equal => (
        "package main; import \"fmt\"; import \"errors\"; var ErrDone = errors.New(\"done\"); func main() { fmt.Println(ErrDone == ErrDone) }",
        vec!["true"]
    ),
    fmt_errorf_formats_without_wrap => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"code %d\", 404); fmt.Println(err.Error()) }",
        vec!["code 404"]
    ),
    fmt_errorf_wrap_includes_cause_text => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := fmt.Errorf(\"read failed: %w\", errors.New(\"EOF\")); fmt.Println(err.Error()) }",
        vec!["read failed: EOF"]
    ),
    fmt_errorf_without_wrap_not_in_chain => (
        "package main; import \"fmt\"; import \"errors\"; var ErrRoot = errors.New(\"root\"); func main() { err := fmt.Errorf(\"outer: %v\", ErrRoot); fmt.Println(errors.Is(err, ErrRoot)) }",
        vec!["false"]
    ),
    errors_is_matches_sentinel => (
        "package main; import \"fmt\"; import \"errors\"; var ErrNotFound = errors.New(\"not found\"); func main() { err := fmt.Errorf(\"open: %w\", ErrNotFound); fmt.Println(errors.Is(err, ErrNotFound)) }",
        vec!["true"]
    ),
    errors_is_rejects_unrelated_target => (
        "package main; import \"fmt\"; import \"errors\"; var ErrA = errors.New(\"a\"); var ErrB = errors.New(\"b\"); func main() { fmt.Println(errors.Is(ErrA, ErrB)) }",
        vec!["false"]
    ),
    errors_is_nil_target_always_false => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.Is(errors.New(\"x\"), nil)) }",
        vec!["false"]
    ),
    errors_is_nil_err_false => (
        "package main; import \"fmt\"; import \"errors\"; var ErrX = errors.New(\"x\"); func main() { fmt.Println(errors.Is(nil, ErrX)) }",
        vec!["false"]
    ),
    errors_is_through_double_wrap => (
        "package main; import \"fmt\"; import \"errors\"; var ErrBase = errors.New(\"base\"); func main() { err := fmt.Errorf(\"layer1: %w\", fmt.Errorf(\"layer2: %w\", ErrBase)); fmt.Println(errors.Is(err, ErrBase)) }",
        vec!["true"]
    ),
    errors_wrapped_not_equal_but_is_true => (
        "package main; import \"fmt\"; import \"errors\"; var ErrSentinel = errors.New(\"sentinel\"); func main() { wrapped := fmt.Errorf(\"wrap: %w\", ErrSentinel); fmt.Println(wrapped == ErrSentinel); fmt.Println(errors.Is(wrapped, ErrSentinel)) }",
        vec!["false", "true"]
    ),
    errors_as_finds_custom_type => (
        "package main; import \"fmt\"; import \"errors\"; type coded struct { n int }; func (c coded) Error() string { return \"coded\" }; func main() { err := error(coded{n: 7}); var target coded; fmt.Println(errors.As(err, &target)); fmt.Println(target.n) }",
        vec!["true", "7"]
    ),
    errors_as_misses_unrelated_type => (
        "package main; import \"fmt\"; import \"errors\"; type coded struct { n int }; func (c coded) Error() string { return \"coded\" }; func main() { err := errors.New(\"plain\"); var target coded; fmt.Println(errors.As(err, &target)) }",
        vec!["false"]
    ),
    errors_as_through_fmt_wrap => (
        "package main; import \"fmt\"; import \"errors\"; type coded struct { n int }; func (c coded) Error() string { return \"coded\" }; func main() { inner := coded{n: 3}; err := fmt.Errorf(\"wrapped: %w\", inner); var target coded; fmt.Println(errors.As(err, &target)); fmt.Println(target.n) }",
        vec!["true", "3"]
    ),
    errors_unwrap_returns_immediate_inner => (
        "package main; import \"fmt\"; import \"errors\"; func main() { inner := errors.New(\"inner\"); outer := fmt.Errorf(\"outer: %w\", inner); fmt.Println(errors.Unwrap(outer) == inner) }",
        vec!["true"]
    ),
    errors_unwrap_plain_error_nil => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.Unwrap(errors.New(\"plain\")) == nil) }",
        vec!["true"]
    ),
    errors_unwrap_nil_err_nil => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.Unwrap(nil) == nil) }",
        vec!["true"]
    ),
    errors_unwrap_join_returns_nil => (
        "package main; import \"fmt\"; import \"errors\"; func main() { joined := errors.Join(errors.New(\"a\"), errors.New(\"b\")); fmt.Println(errors.Unwrap(joined) == nil) }",
        vec!["true"]
    ),
    errors_join_formats_with_newlines => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.Join(errors.New(\"first\"), errors.New(\"second\")); fmt.Println(err.Error()) }",
        vec!["first\nsecond"]
    ),
    errors_join_all_nil_returns_nil => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.Join(nil, nil) == nil) }",
        vec!["true"]
    ),
    errors_join_single_error_message => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.Join(errors.New(\"solo\")); fmt.Println(err.Error()) }",
        vec!["solo"]
    ),
    errors_join_is_finds_member => (
        "package main; import \"fmt\"; import \"errors\"; func main() { one := errors.New(\"one\"); two := errors.New(\"two\"); joined := errors.Join(one, two); fmt.Println(errors.Is(joined, one)) }",
        vec!["true"]
    ),
    errors_println_default_string => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.New(\"boom\")) }",
        vec!["boom"]
    ),
}

go_compile_cases! {
    errors_join_three_constituents => "package main; import \"errors\"; func main() { _ = errors.Join(errors.New(\"a\"), errors.New(\"b\"), errors.New(\"c\")) }",
    errors_as_pointer_to_interface => "package main; import \"errors\"; type timeout interface { Timeout() bool }; func main() { var target timeout; _ = errors.As(errors.New(\"x\"), &target) }",
    errors_custom_is_method => "package main; import \"errors\"; var ErrSpecial = errors.New(\"special\"); type wrapper struct { err error }; func (w wrapper) Error() string { return w.err.Error() }; func (w wrapper) Unwrap() error { return w.err }; func (w wrapper) Is(target error) bool { return target == ErrSpecial }; func main() { _ = errors.Is(wrapper{err: ErrSpecial}, ErrSpecial) }",
    fmt_errorf_multiple_args_no_wrap => "package main; import \"fmt\"; func main() { _ = fmt.Errorf(\"%s %d\", \"err\", 1) }",
    errors_join_filters_nil_entries => "package main; import \"errors\"; func main() { _ = errors.Join(nil, errors.New(\"only\"), nil) }",
    errors_new_empty_message => "package main; import \"errors\"; func main() { _ = errors.New(\"\") }",
    errors_is_on_custom_unwrap_chain => "package main; import \"errors\"; type link struct { next error }; func (l link) Error() string { return \"link\" }; func (l link) Unwrap() error { return l.next }; func main() { base := errors.New(\"base\"); _ = errors.Is(link{next: base}, base) }",
    fmt_errorf_wrap_nil_error => "package main; import \"fmt\"; func main() { _ = fmt.Errorf(\"wrap: %w\", nil) }",
}
