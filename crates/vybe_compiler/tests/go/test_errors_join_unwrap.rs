//! errors.Join, Unwrap, Is/As on joined errors, New/Errorf formatting verbs,
//! sentinel comparison — distinct from `test_errors_package.rs` (basic chains) and
//! `test_fmt_errors_print.rs` (Sscanf/Fprint, non-chain Errorf).


go_run_cases! {
    errors_join_two_messages => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.Join(errors.New(\"alpha\"), errors.New(\"beta\")); fmt.Println(err.Error()) }",
        vec!["alpha\nbeta"]
    ),
    errors_join_three_messages => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.Join(errors.New(\"one\"), errors.New(\"two\"), errors.New(\"three\")); fmt.Println(err.Error()) }",
        vec!["one\ntwo\nthree"]
    ),
    errors_join_single_no_newline => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.Join(errors.New(\"only\")); fmt.Println(err.Error()) }",
        vec!["only"]
    ),
    errors_join_all_nil => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.Join(nil, nil, nil) == nil) }",
        vec!["true"]
    ),
    errors_join_filters_nil_middle => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.Join(errors.New(\"first\"), nil, errors.New(\"last\")); fmt.Println(err.Error()) }",
        vec!["first\nlast"]
    ),
    errors_join_filters_nil_edges => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.Join(nil, errors.New(\"middle\"), nil); fmt.Println(err.Error()) }",
        vec!["middle"]
    ),
    errors_join_is_first_member => (
        "package main; import \"fmt\"; import \"errors\"; func main() { a := errors.New(\"a\"); b := errors.New(\"b\"); joined := errors.Join(a, b); fmt.Println(errors.Is(joined, a)) }",
        vec!["true"]
    ),
    errors_join_is_second_member => (
        "package main; import \"fmt\"; import \"errors\"; func main() { a := errors.New(\"a\"); b := errors.New(\"b\"); joined := errors.Join(a, b); fmt.Println(errors.Is(joined, b)) }",
        vec!["true"]
    ),
    errors_join_is_unrelated_false => (
        "package main; import \"fmt\"; import \"errors\"; func main() { joined := errors.Join(errors.New(\"x\"), errors.New(\"y\")); fmt.Println(errors.Is(joined, errors.New(\"z\"))) }",
        vec!["false"]
    ),
    errors_join_unwrap_returns_nil => (
        "package main; import \"fmt\"; import \"errors\"; func main() { joined := errors.Join(errors.New(\"a\"), errors.New(\"b\")); fmt.Println(errors.Unwrap(joined) == nil) }",
        vec!["true"]
    ),
    errors_join_not_nil => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.Join(errors.New(\"fail\")); fmt.Println(err != nil) }",
        vec!["true"]
    ),

    errors_unwrap_single_wrap => (
        "package main; import \"fmt\"; import \"errors\"; func main() { inner := errors.New(\"core\"); outer := fmt.Errorf(\"wrap: %w\", inner); fmt.Println(errors.Unwrap(outer).Error()) }",
        vec!["core"]
    ),
    errors_unwrap_double_wrap => (
        "package main; import \"fmt\"; import \"errors\"; func main() { base := errors.New(\"base\"); mid := fmt.Errorf(\"mid: %w\", base); outer := fmt.Errorf(\"outer: %w\", mid); fmt.Println(errors.Unwrap(outer).Error()) }",
        vec!["mid: base"]
    ),
    errors_unwrap_plain_nil => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.Unwrap(errors.New(\"plain\")) == nil) }",
        vec!["true"]
    ),
    errors_unwrap_nil_error => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.Unwrap(nil) == nil) }",
        vec!["true"]
    ),
    errors_unwrap_chain_three_levels => (
        "package main; import \"fmt\"; import \"errors\"; var ErrLeaf = errors.New(\"leaf\"); func main() { e1 := fmt.Errorf(\"l1: %w\", ErrLeaf); e2 := fmt.Errorf(\"l2: %w\", e1); fmt.Println(errors.Is(errors.Unwrap(e2), e1)) }",
        vec!["true"]
    ),

    errors_is_joined_sentinel_first => (
        "package main; import \"fmt\"; import \"errors\"; var ErrOne = errors.New(\"one\"); func main() { joined := errors.Join(ErrOne, errors.New(\"two\")); fmt.Println(errors.Is(joined, ErrOne)) }",
        vec!["true"]
    ),
    errors_is_joined_sentinel_second => (
        "package main; import \"fmt\"; import \"errors\"; var ErrTwo = errors.New(\"two\"); func main() { joined := errors.Join(errors.New(\"one\"), ErrTwo); fmt.Println(errors.Is(joined, ErrTwo)) }",
        vec!["true"]
    ),
    errors_is_joined_wrapped_sentinel => (
        "package main; import \"fmt\"; import \"errors\"; var ErrBase = errors.New(\"base\"); func main() { wrapped := fmt.Errorf(\"w: %w\", ErrBase); joined := errors.Join(wrapped, errors.New(\"other\")); fmt.Println(errors.Is(joined, ErrBase)) }",
        vec!["true"]
    ),
    errors_as_joined_custom_type => (
        "package main; import \"fmt\"; import \"errors\"; type coded struct { code int }; func (c coded) Error() string { return fmt.Sprintf(\"code %d\", c.code) }; func main() { inner := coded{code: 42}; joined := errors.Join(inner, errors.New(\"plain\")); var target coded; fmt.Println(errors.As(joined, &target)); fmt.Println(target.code) }",
        vec!["true", "42"]
    ),
    errors_as_joined_misses_second => (
        "package main; import \"fmt\"; import \"errors\"; type coded struct { code int }; func (c coded) Error() string { return \"coded\" }; func main() { joined := errors.Join(errors.New(\"plain\"), coded{code: 1}); var target coded; fmt.Println(errors.As(joined, &target)) }",
        vec!["false"]
    ),
    errors_is_joined_nil_target => (
        "package main; import \"fmt\"; import \"errors\"; func main() { joined := errors.Join(errors.New(\"x\")); fmt.Println(errors.Is(joined, nil)) }",
        vec!["false"]
    ),

    errors_new_empty_string => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.New(\"\").Error()) }",
        vec![""]
    ),
    errors_new_multiline => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.New(\"line1\\nline2\").Error()) }",
        vec!["line1\nline2"]
    ),
    errors_new_unicode => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.New(\"エラー\").Error()) }",
        vec!["エラー"]
    ),

    errorf_percent_s_verb => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"host %s down\", \"api\"); fmt.Println(err.Error()) }",
        vec!["host api down"]
    ),
    errorf_percent_d_verb => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"exit code %d\", 127); fmt.Println(err.Error()) }",
        vec!["exit code 127"]
    ),
    errorf_percent_v_verb => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"data %v\", []int{1, 2}); fmt.Println(err.Error()) }",
        vec!["data [1 2]"]
    ),
    errorf_percent_plus_v_verb => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"detail %+v\", struct{ ID int }{ID: 9}); fmt.Println(err.Error()) }",
        vec!["detail {ID: 9}"]
    ),
    errorf_percent_q_verb => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"token %q\", \"tab\\there\"); fmt.Println(err.Error()) }",
        vec!["token \"tab\\there\""]
    ),
    errorf_percent_x_verb => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"addr 0x%x\", 255); fmt.Println(err.Error()) }",
        vec!["addr 0xff"]
    ),
    errorf_percent_t_verb => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"type %T\", 42); fmt.Println(err.Error()) }",
        vec!["type int"]
    ),
    errorf_wrap_preserves_chain => (
        "package main; import \"fmt\"; import \"errors\"; var ErrIO = errors.New(\"io\"); func main() { err := fmt.Errorf(\"read: %w\", ErrIO); fmt.Println(err.Error()); fmt.Println(errors.Is(err, ErrIO)) }",
        vec!["read: io", "true"]
    ),
    errorf_multiple_wraps => (
        "package main; import \"fmt\"; import \"errors\"; func main() { e1 := fmt.Errorf(\"a: %w\", errors.New(\"b\")); e2 := fmt.Errorf(\"c: %w\", e1); fmt.Println(errors.Is(e2, errors.New(\"b\"))) }",
        vec!["false"]
    ),
    errorf_wrap_with_formatting => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := fmt.Errorf(\"failed after %d retries: %w\", 3, errors.New(\"timeout\")); fmt.Println(err.Error()) }",
        vec!["failed after 3 retries: timeout"]
    ),

    sentinel_same_reference_equal => (
        "package main; import \"fmt\"; import \"errors\"; var ErrEOF = errors.New(\"EOF\"); func main() { fmt.Println(ErrEOF == ErrEOF) }",
        vec!["true"]
    ),
    sentinel_different_vars_same_text_unequal => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.New(\"EOF\") == errors.New(\"EOF\")) }",
        vec!["false"]
    ),
    sentinel_package_level_equal => (
        "package main; import \"fmt\"; import \"errors\"; var ErrA = errors.New(\"fail\"); var ErrB = ErrA; func main() { fmt.Println(ErrA == ErrB) }",
        vec!["true"]
    ),
    sentinel_is_not_equality => (
        "package main; import \"fmt\"; import \"errors\"; var ErrSent = errors.New(\"sent\"); func main() { other := fmt.Errorf(\"wrap: %w\", ErrSent); fmt.Println(other == ErrSent); fmt.Println(errors.Is(other, ErrSent)) }",
        vec!["false", "true"]
    ),
    sentinel_joined_same_var => (
        "package main; import \"fmt\"; import \"errors\"; var ErrX = errors.New(\"x\"); func main() { joined := errors.Join(ErrX, errors.New(\"y\")); fmt.Println(errors.Is(joined, ErrX)) }",
        vec!["true"]
    ),
    sentinel_compare_to_nil => (
        "package main; import \"fmt\"; import \"errors\"; var ErrZ = errors.New(\"z\"); func main() { fmt.Println(ErrZ == nil) }",
        vec!["false"]
    ),
    sentinel_wrapped_not_equal_unwrapped => (
        "package main; import \"fmt\"; import \"errors\"; var ErrRoot = errors.New(\"root\"); func main() { w := fmt.Errorf(\"w: %w\", ErrRoot); fmt.Println(w == ErrRoot) }",
        vec!["false"]
    ),

    errors_join_four_constituents => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.Join(errors.New(\"a\"), errors.New(\"b\"), errors.New(\"c\"), errors.New(\"d\")); parts := 0; for _, ch := range err.Error() { if ch == '\\n' { parts++ } }; fmt.Println(parts) }",
        vec!["3"]
    ),
    errors_join_empty_strings => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.Join(errors.New(\"\"), errors.New(\"\")); fmt.Println(len(err.Error())) }",
        vec!["1"]
    ),
    errors_is_on_joined_wrapped_chain => (
        "package main; import \"fmt\"; import \"errors\"; var ErrDeep = errors.New(\"deep\"); func main() { wrapped := fmt.Errorf(\"layer: %w\", ErrDeep); joined := errors.Join(errors.New(\"shallow\"), wrapped); fmt.Println(errors.Is(joined, ErrDeep)) }",
        vec!["true"]
    ),
    errors_as_on_wrapped_in_join => (
        "package main; import \"fmt\"; import \"errors\"; type myErr struct { msg string }; func (e myErr) Error() string { return e.msg }; func main() { inner := myErr{msg: \"inner\"}; joined := errors.Join(fmt.Errorf(\"wrap: %w\", inner)); var target myErr; fmt.Println(errors.As(joined, &target)); fmt.Println(target.msg) }",
        vec!["true", "inner"]
    ),
    errorf_bool_verb => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"ok=%t fail=%t\", true, false); fmt.Println(err.Error()) }",
        vec!["ok=true fail=false"]
    ),
    errorf_float_verb => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"rate %.1f%%\", 99.5); fmt.Println(err.Error()) }",
        vec!["rate 99.5%"]
    ),
    errorf_pointer_verb => (
        "package main; import \"fmt\"; func main() { n := 7; err := fmt.Errorf(\"ptr %p\", &n); fmt.Println(len(err.Error()) > 0) }",
        vec!["true"]
    ),
    errors_new_distinct_from_fmt_errorf => (
        "package main; import \"fmt\"; import \"errors\"; func main() { a := errors.New(\"msg\"); b := fmt.Errorf(\"msg\"); fmt.Println(a == b); fmt.Println(a.Error() == b.Error()) }",
        vec!["false", "true"]
    ),
    errors_join_with_sentinel_and_plain => (
        "package main; import \"fmt\"; import \"errors\"; var ErrFatal = errors.New(\"fatal\"); func main() { err := errors.Join(ErrFatal, errors.New(\"warning\")); fmt.Println(errors.Is(err, ErrFatal)); fmt.Println(errors.Is(err, errors.New(\"warning\"))) }",
        vec!["true", "false"]
    ),
    errors_unwrap_joined_not_chain => (
        "package main; import \"fmt\"; import \"errors\"; func main() { inner := errors.New(\"inner\"); joined := errors.Join(fmt.Errorf(\"wrap: %w\", inner)); fmt.Println(errors.Is(joined, inner)) }",
        vec!["true"]
    ),
    errors_is_self => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.New(\"self\"); fmt.Println(errors.Is(err, err)) }",
        vec!["true"]
    ),
    errors_as_nil_target => (
        "package main; import \"fmt\"; import \"errors\"; func main() { fmt.Println(errors.As(errors.New(\"x\"), nil)) }",
        vec!["false"]
    ),
    sentinel_reassign_same => (
        "package main; import \"fmt\"; import \"errors\"; var ErrOld = errors.New(\"old\"); func main() { ErrOld = errors.New(\"new\"); fmt.Println(ErrOld.Error()) }",
        vec!["new"]
    ),
    errorf_width_padding => (
        "package main; import \"fmt\"; func main() { err := fmt.Errorf(\"code %05d\", 7); fmt.Println(err.Error()) }",
        vec!["code 00007"]
    ),
    errorf_stringer_verb => (
        "package main; import \"fmt\"; type id int; func (i id) String() string { return fmt.Sprintf(\"ID-%d\", i) }; func main() { err := fmt.Errorf(\"entity %s\", id(5)); fmt.Println(err.Error()) }",
        vec!["entity ID-5"]
    ),
    errors_join_one_nil_one_error => (
        "package main; import \"fmt\"; import \"errors\"; func main() { err := errors.Join(nil, errors.New(\"solo\")); fmt.Println(err.Error()) }",
        vec!["solo"]
    ),
    errors_is_wrapped_sentinel_deep => (
        "package main; import \"fmt\"; import \"errors\"; var ErrBottom = errors.New(\"bottom\"); func main() { e := fmt.Errorf(\"a: %w\", fmt.Errorf(\"b: %w\", fmt.Errorf(\"c: %w\", ErrBottom))); fmt.Println(errors.Is(e, ErrBottom)) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    errors_join_many => "package main; import \"errors\"; func main() { _ = errors.Join(errors.New(\"a\"), errors.New(\"b\"), errors.New(\"c\"), errors.New(\"d\"), errors.New(\"e\")) }",
    errors_as_interface_target => "package main; import \"errors\"; type timeout interface { Timeout() bool }; func main() { var t timeout; _ = errors.As(errors.New(\"x\"), &t) }",
    errors_custom_is_on_join => "package main; import \"errors\"; var ErrSpec = errors.New(\"spec\"); type w struct { e error }; func (w w) Error() string { return w.e.Error() }; func (w w) Is(t error) bool { return t == ErrSpec }; func main() { _ = errors.Is(w{e: ErrSpec}, ErrSpec) }",
    errorf_wrap_nil => "package main; import \"fmt\"; func main() { _ = fmt.Errorf(\"wrap: %w\", nil) }",
    errorf_multiple_args => "package main; import \"fmt\"; func main() { _ = fmt.Errorf(\"%s %d %v\", \"err\", 1, true) }",
    errors_unwrap_custom => "package main; import \"errors\"; type chain struct { next error }; func (c chain) Error() string { return \"chain\" }; func (c chain) Unwrap() error { return c.next }; func main() { _ = errors.Unwrap(chain{next: errors.New(\"inner\")}) }",
    errors_join_with_fmt_wrap => "package main; import \"fmt\"; import \"errors\"; func main() { _ = errors.Join(fmt.Errorf(\"w: %w\", errors.New(\"inner\"))) }",
    errors_is_nil_err => "package main; import \"errors\"; var ErrT = errors.New(\"t\"); func main() { _ = errors.Is(nil, ErrT) }",
    errorf_complex_verbs => "package main; import \"fmt\"; func main() { _ = fmt.Errorf(\"%#v %+#v\", 1, struct{ X int }{1}) }",
    errors_new_long_message => "package main; import \"errors\"; func main() { _ = errors.New(\"a long error message for compile smoke\") }",
    errors_as_struct_pointer => "package main; import \"errors\"; type E struct { N int }; func (e E) Error() string { return \"e\" }; func main() { var target E; _ = errors.As(E{N: 1}, &target) }",
    errorf_percent_w_only => "package main; import \"fmt\"; import \"errors\"; func main() { _ = fmt.Errorf(\"%w\", errors.New(\"cause\")) }",
    errors_join_sentinel_var => "package main; import \"errors\"; var ErrJ = errors.New(\"j\"); func main() { _ = errors.Join(ErrJ, ErrJ) }",
    errors_unwrap_multi_no_unwrap => "package main; import \"errors\"; type plain struct{}; func (plain) Error() string { return \"p\" }; func main() { _ = errors.Unwrap(plain{}) }",
    errorf_rune_verb => "package main; import \"fmt\"; func main() { _ = fmt.Errorf(\"char %c\", 'A') }",
}
