//! Advanced function types: distinct named signatures, returning functions,
//! nil comparison, struct fields, and methods with function parameters.

go_run_cases! {
    // --- nil comparison on function values ---
    two_zero_value_funcs_same_type_equal =>
        ("package main; import \"fmt\"; func main() { var left func(); var right func(); fmt.Println(left == right); fmt.Println(left == nil) }",
        vec!["true", "true"]),
    assigned_func_literal_not_nil =>
        ("package main; import \"fmt\"; func main() { var fn func(int) int; fmt.Println(fn == nil); fn = func(v int) int { return v + 1 }; fmt.Println(fn == nil) }",
        vec!["true", "false"]),
    func_with_params_cleared_to_nil =>
        ("package main; import \"fmt\"; func main() { var fn func(int, int) int; fn = func(a int, b int) int { return a + b }; fmt.Println(fn(2, 3)); fn = nil; fmt.Println(fn == nil) }",
        vec!["5", "true"]),
    func_returning_func_zero_value_nil =>
        ("package main; import \"fmt\"; func main() { var factory func(int) func(int) int; fmt.Println(factory == nil) }",
        vec!["true"]),
    two_nil_funcs_different_signatures_both_nil =>
        ("package main; import \"fmt\"; func main() { var noop func(); var add func(int) int; fmt.Println(noop == nil); fmt.Println(add == nil) }",
        vec!["true", "true"]),

    // --- distinct named function types as parameters ---
    named_mapper_type_passed_and_called =>
        ("package main; import \"fmt\"; type Mapper func(int) int; func apply(v int, m Mapper) int { return m(v) }; func main() { fmt.Println(apply(4, Mapper(func(v int) int { return v * 3 }))) }",
        vec!["12"]),
    named_handler_type_string_result =>
        ("package main; import \"fmt\"; type Handler func() string; func run(h Handler) string { return h() }; func main() { fmt.Println(run(Handler(func() string { return \"ok\" }))) }",
        vec!["ok"]),
    named_reducer_type_two_param_callback =>
        ("package main; import \"fmt\"; type Reducer func(int, int) int; func fold(values []int, r Reducer, init int) int { acc := init; for _, v := range values { acc = r(acc, v) }; return acc }; func main() { fmt.Println(fold([]int{1, 2, 3}, Reducer(func(a int, b int) int { return a + b }), 0)) }",
        vec!["6"]),
    distinct_named_types_same_signature_explicit_cast =>
        ("package main; import \"fmt\"; type Adder func(int, int) int; type Combiner func(int, int) int; func use(c Combiner) int { return c(2, 5) }; func main() { var add Adder = func(a int, b int) int { return a + b }; fmt.Println(use(Combiner(add))) }",
        vec!["7"]),

    // --- returning functions ---
    return_curried_multiplier_applied_twice =>
        ("package main; import \"fmt\"; func scale(factor int) func(int) int { return func(v int) int { return v * factor } }; func main() { double := scale(2); triple := scale(3); fmt.Println(double(4)); fmt.Println(triple(4)) }",
        vec!["8", "12"]),
    return_func_selected_by_boolean_flag =>
        ("package main; import \"fmt\"; func pick(positive bool) func(int) int { if positive { return func(v int) int { return v + 1 } }; return func(v int) int { return v - 1 } }; func main() { fmt.Println(pick(true)(6)); fmt.Println(pick(false)(6)) }",
        vec!["7", "5"]),
    return_func_fallback_when_primary_nil =>
        ("package main; import \"fmt\"; func fallback(primary func() int, backup func() int) int { if primary != nil { return primary() }; return backup() }; func main() { fmt.Println(fallback(nil, func() int { return 42 })) }",
        vec!["42"]),
    return_explicit_nil_func_when_disabled =>
        ("package main; import \"fmt\"; func maybe(enabled bool) func() { if enabled { return func() {} }; return nil }; func main() { fmt.Println(maybe(false) == nil); fmt.Println(maybe(true) == nil) }",
        vec!["true", "false"]),
    returned_func_captures_outer_by_reference =>
        ("package main; import \"fmt\"; func counter() func() int { total := 0; return func() int { total++; return total } }; func main() { next := counter(); fmt.Println(next()); fmt.Println(next()) }",
        vec!["1", "2"]),

    // --- struct fields holding function values ---
    struct_func_field_invoked_directly =>
        ("package main; import \"fmt\"; type worker struct { run func(int) int }; func main() { value := worker{run: func(v int) int { return v + 5 }}; fmt.Println(value.run(7)) }",
        vec!["12"]),
    struct_two_func_fields_distinct_results =>
        ("package main; import \"fmt\"; type pair struct { left func() int; right func() int }; func main() { value := pair{left: func() int { return 1 }, right: func() int { return 2 }}; fmt.Println(value.left()); fmt.Println(value.right()) }",
        vec!["1", "2"]),
    struct_func_field_reassigned_and_called =>
        ("package main; import \"fmt\"; type holder struct { fn func(string) string }; func main() { value := holder{}; fmt.Println(value.fn == nil); value.fn = func(s string) string { return s + \"!\" }; fmt.Println(value.fn(\"go\")) }",
        vec!["true", "go!"]),
    struct_field_func_cast_to_named_type =>
        ("package main; import \"fmt\"; type Mapper func(int) int; type holder struct { fn func(int) int }; func apply(v int, m Mapper) int { return m(v) }; func main() { value := holder{fn: func(v int) int { return v + 2 }}; fmt.Println(apply(5, Mapper(value.fn))) }",
        vec!["7"]),
    struct_passes_field_func_to_helper =>
        ("package main; import \"fmt\"; type box struct { transform func(int) int }; func invoke(b box, v int) int { return b.transform(v) }; func main() { value := box{transform: func(v int) int { return v - 1 }}; fmt.Println(invoke(value, 9)) }",
        vec!["8"]),

    // --- methods with function parameters ---
    method_runs_supplied_transform_on_field =>
        ("package main; import \"fmt\"; type gauge struct { value int }; func (g *gauge) mapValue(mapper func(int) int) { g.value = mapper(g.value) }; func main() { g := gauge{value: 4}; g.mapValue(func(v int) int { return v * 2 }); fmt.Println(g.value) }",
        vec!["8"]),
    method_apply_twice_with_predicate_param =>
        ("package main; import \"fmt\"; type tally struct { count int }; func (t *tally) whilePositive(ok func(int) bool) { for ok(t.count) { t.count-- } }; func main() { value := tally{count: 3}; value.whilePositive(func(v int) bool { return v > 0 }); fmt.Println(value.count) }",
        vec!["0"]),
    method_for_each_with_index_callback =>
        ("package main; import \"fmt\"; type batch struct { items []int }; func (b batch) forEach(visit func(int, int)) { for i, v := range b.items { visit(i, v) } }; func main() { sum := 0; batch{items: []int{2, 3, 4}}.forEach(func(i int, v int) { sum += v }); fmt.Println(sum) }",
        vec!["9"]),
    pointer_receiver_run_with_func_param =>
        ("package main; import \"fmt\"; type acc struct { total int }; func (a *acc) addEach(values []int, combine func(int, int) int) { for _, v := range values { a.total = combine(a.total, v) } }; func main() { value := acc{}; value.addEach([]int{1, 2, 3}, func(a int, b int) int { return a + b }); fmt.Println(value.total) }",
        vec!["6"]),
    method_passes_own_func_field_to_visitor =>
        ("package main; import \"fmt\"; type node struct { label string; format func(string) string }; func (n node) show(visitor func(string)) { visitor(n.format(n.label)) }; func main() { value := node{label: \"go\", format: func(s string) string { return s + \"!\" }}; value.show(func(s string) { fmt.Println(s) }) }",
        vec!["go!"]),
}

go_compile_cases! {
  distinct_named_func_types_same_signature_compile =>
    "package main; type Step func(int) int; type Stage func(int) int; func pipe(v int, s Stage) int { return s(v) }; func main() { var step Step = func(v int) int { return v + 1 }; _ = pipe(1, Stage(step)) }",
  func_param_and_return_distinct_signatures_compile =>
    "package main; func read(fn func() int) int { return fn() }; func write(fn func(int)) {}; func main() { write(func(v int) {}); _ = read(func() int { return 1 }) }",
  variadic_func_type_distinct_from_fixed_compile =>
    "package main; type Fixed func(int) int; type Variadic func(...int) int; func useFixed(f Fixed) int { return f(1) }; func useVariadic(v Variadic) int { return v(1, 2) }; func main() { _ = useFixed(func(v int) int { return v }); _ = useVariadic(func(values ...int) int { return len(values) }) }",
  struct_field_func_type_with_named_alias_compile =>
    "package main; type Callback func(int) bool; type registry struct { filter Callback }; func main() { _ = registry{filter: Callback(func(v int) bool { return v > 0 })} }",
  method_accepts_named_func_type_param_compile =>
    "package main; type Predicate func(int) bool; type sieve struct { n int }; func (s sieve) keep(p Predicate) bool { return p(s.n) }; func main() { _ = sieve{n: 2}.keep(Predicate(func(v int) bool { return v%2 == 0 })) }",
  nested_func_return_type_as_field_compile =>
    "package main; type Factory func() func() int; type holder struct { build Factory }; func main() { _ = holder{build: Factory(func() func() int { return func() int { return 1 } })} }",
  slice_of_named_func_type_compile =>
    "package main; type Op func(int) int; func main() { ops := []Op{Op(func(v int) int { return v }), Op(func(v int) int { return v + 1 })}; _ = ops[1](2) }",
  method_returns_closure_from_receiver_compile =>
    "package main; type builder struct { base int }; func (b builder) incrementer() func(int) int { return func(v int) int { return b.base + v } }; func main() { _ = builder{base: 10}.incrementer() }",
    method_func_param_with_multiple_returns_compile =>
    "package main; type Splitter func(int) (int, int); type divider struct{}; func (divider) parts(v int, split Splitter) (int, int) { return split(v) }; func main() { _, _ = divider{}.parts(9, Splitter(func(v int) (int, int) { return v / 2, v % 2 })) }",
}
