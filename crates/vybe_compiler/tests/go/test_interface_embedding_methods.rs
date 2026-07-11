//! Interface embedding — promoted method dispatch, overlapping method sets, and nil
//! interface method calls. Distinct from `test_interfaces_patterns_extra.rs` (general
//! interface patterns) and `test_struct_embedding_advanced.rs` (struct field/method promotion).

go_run_cases! {
    reader_writer_read_promoted_runtime =>
        ("package main; import \"fmt\"; type reader interface { read() int }; type writer interface { write(int) }; type readWriter interface { reader; writer }; type buf struct { data int }; func (b buf) read() int { return b.data }; func (b buf) write(n int) { b.data = n }; func main() { var rw readWriter = buf{data: 7}; fmt.Println(rw.read()) }", vec!["7"]),
    reader_writer_write_promoted_runtime =>
        ("package main; import \"fmt\"; type reader interface { read() int }; type writer interface { write(int) }; type readWriter interface { reader; writer }; type buf struct { data int }; func (b *buf) read() int { return b.data }; func (b *buf) write(n int) { b.data = n }; func main() { value := &buf{}; var rw readWriter = value; rw.write(9); fmt.Println(rw.read()) }", vec!["9"]),
    triple_embedded_interface_dispatch_runtime =>
        ("package main; import \"fmt\"; type leaf interface { tag() string }; type branch interface { leaf }; type trunk interface { branch }; type node struct{}; func (node) tag() string { return \"deep\" }; func main() { var t trunk = node{}; fmt.Println(t.tag()) }", vec!["deep"]),
    composite_interface_struct_impl_runtime =>
        ("package main; import \"fmt\"; type opener interface { open() bool }; type closer interface { close() }; type resource interface { opener; closer }; type file struct { ok bool }; func (f file) open() bool { return f.ok }; func (f file) close() {}; func main() { var r resource = file{ok: true}; fmt.Println(r.open()) }", vec!["true"]),
    promoted_method_with_int_arg_runtime =>
        ("package main; import \"fmt\"; type counter interface { bump(int) int }; type meter interface { counter }; type gauge struct { n int }; func (g gauge) bump(delta int) int { return g.n + delta }; func main() { var m meter = gauge{n: 4}; fmt.Println(m.bump(3)) }", vec!["7"]),
    promoted_method_returns_string_runtime =>
        ("package main; import \"fmt\"; type namer interface { name() string }; type labeled interface { namer }; type widget struct{}; func (widget) name() string { return \"vybe\" }; func main() { var value labeled = widget{}; fmt.Println(value.name()) }", vec!["vybe"]),
    three_embedded_distinct_methods_runtime =>
        ("package main; import \"fmt\"; type alpha interface { a() int }; type beta interface { b() int }; type gamma interface { c() int }; type combo interface { alpha; beta; gamma }; type triple struct{}; func (triple) a() int { return 1 }; func (triple) b() int { return 2 }; func (triple) c() int { return 3 }; func main() { var value combo = triple{}; fmt.Println(value.a()); fmt.Println(value.b()); fmt.Println(value.c()) }", vec!["1", "2", "3"]),
    overlapping_method_unified_impl_runtime =>
        ("package main; import \"fmt\"; type resetterA interface { reset() int }; type resetterB interface { reset() int }; type dualReset interface { resetterA; resetterB }; type engine struct { ticks int }; func (e *engine) reset() int { e.ticks = 0; return e.ticks }; func main() { value := &engine{ticks: 5}; var d dualReset = value; fmt.Println(d.reset()) }", vec!["0"]),
    composite_interface_as_function_arg_runtime =>
        ("package main; import \"fmt\"; type speaker interface { speak() string }; type louder interface { speaker }; type dog struct{}; func (dog) speak() string { return \"woof\" }; func echo(value louder) string { return value.speak() }; func main() { fmt.Println(echo(dog{})) }", vec!["woof"]),
    composite_interface_in_slice_runtime =>
        ("package main; import \"fmt\"; type mover interface { move() int }; type jumper interface { mover }; type hopper struct { steps int }; func (h hopper) move() int { return h.steps }; func main() { values := []jumper{hopper{steps: 3}}; fmt.Println(values[0].move()) }", vec!["3"]),
    composite_interface_reassign_implementer_runtime =>
        ("package main; import \"fmt\"; type sized interface { size() int }; type measurable interface { sized }; type small struct{}; func (small) size() int { return 1 }; type large struct{}; func (large) size() int { return 9 }; func main() { var m measurable = small{}; fmt.Println(m.size()); m = large{}; fmt.Println(m.size()) }", vec!["1", "9"]),
    promoted_interface_method_value_runtime =>
        ("package main; import \"fmt\"; type greeter interface { greet() string }; type social interface { greeter }; type hi struct{}; func (hi) greet() string { return \"hi\" }; func main() { var s social = hi{}; fn := s.greet; fmt.Println(fn()) }", vec!["hi"]),
    pointer_receiver_embedded_interface_runtime =>
        ("package main; import \"fmt\"; type scaler interface { scale() int }; type resizable interface { scaler }; type icon struct { n int }; func (i *icon) scale() int { return i.n * 2 }; func main() { var r resizable = &icon{n: 5}; fmt.Println(r.scale()) }", vec!["10"]),
    value_receiver_embedded_interface_runtime =>
        ("package main; import \"fmt\"; type scaler interface { scale() int }; type resizable interface { scaler }; type icon struct { n int }; func (i icon) scale() int { return i.n * 2 }; func main() { var r resizable = icon{n: 6}; fmt.Println(r.scale()) }", vec!["12"]),
    four_level_interface_embed_dispatch_runtime =>
        ("package main; import \"fmt\"; type d interface { n() int }; type c interface { d }; type b interface { c }; type a interface { b }; type leaf struct { value int }; func (l leaf) n() int { return l.value }; func main() { var top a = leaf{value: 13}; fmt.Println(top.n()) }", vec!["13"]),
    left_right_embedded_interface_methods_runtime =>
        ("package main; import \"fmt\"; type left interface { side() string }; type right interface { edge() string }; type pair interface { left; right }; type both struct{}; func (both) side() string { return \"L\" }; func (both) edge() string { return \"R\" }; func main() { var p pair = both{}; fmt.Println(p.side()); fmt.Println(p.edge()) }", vec!["L", "R"]),
    composite_interface_struct_field_call_runtime =>
        ("package main; import \"fmt\"; type runner interface { run() int }; type athlete interface { runner }; type team struct { lead athlete }; type sprinter struct { pace int }; func (s sprinter) run() int { return s.pace }; func main() { squad := team{lead: sprinter{pace: 42}}; fmt.Println(squad.lead.run()) }", vec!["42"]),
    embedded_interface_method_chain_runtime =>
        ("package main; import \"fmt\"; type builder interface { build() string }; type factory interface { builder }; type widget struct{}; func (widget) build() string { return \"built\" }; func makeFactory() factory { return widget{} }; func main() { fmt.Println(makeFactory().build()) }", vec!["built"]),
}

go_compile_cases! {
    overlapping_identical_method_two_embed_compile =>
        "package main; type resetterA interface { reset() }; type resetterB interface { reset() }; type dualReset interface { resetterA; resetterB }; type engine struct{}; func (engine) reset() {}; func main() { var value dualReset = engine{}; value.reset() }",
    overlapping_method_three_interfaces_compile =>
        "package main; type a interface { ping() int }; type b interface { ping() int }; type c interface { ping() int }; type trio interface { a; b; c }; type echo struct{}; func (echo) ping() int { return 1 }; func main() { var value trio = echo{}; _ = value.ping() }",
    nil_composite_interface_guarded_call_compile =>
        "package main; type worker interface { work() }; type task interface { worker }; func run(value task) { if value != nil { value.work() } }; func main() { run(nil) }",
    nil_composite_interface_unchecked_call_compile =>
        "package main; type worker interface { work() }; type job interface { worker }; func main() { var value job; value.work() }",
    nil_composite_passed_to_callee_compile =>
        "package main; type speaker interface { speak() string }; type talker interface { speaker }; func say(value talker) string { if value == nil { return \"nil\" }; return value.speak() }; func main() { _ = say(nil) }",
    nil_composite_method_in_defer_compile =>
        "package main; type cleaner interface { clean() }; type janitor interface { cleaner }; func sweep(value janitor) { defer value.clean() }; func main() { sweep(nil) }",
    triple_interface_embed_definition_compile =>
        "package main; type leaf interface { tag() string }; type branch interface { leaf }; type trunk interface { branch }; func main() {}",
    composite_interface_method_expression_compile =>
        "package main; type mover interface { move() int }; type walker interface { mover }; type step func() int; func (s step) move() int { return s() }; func main() { var fn func(walker) func() int = walker.move; _ = fn }",
    composite_satisfies_via_struct_literal_compile =>
        "package main; type lock interface { lock() }; type unlock interface { unlock() }; type mutex interface { lock; unlock }; type gate struct{}; func (gate) lock() {}; func (gate) unlock() {}; func main() { var m mutex = gate{}; m.lock() }",
    overlapping_embedded_seeker_teller_compile =>
        "package main; type seeker interface { tell() int }; type pointer interface { tell() int }; type locator interface { seeker; pointer }; type cursor struct { pos int }; func (c cursor) tell() int { return c.pos }; func main() { var loc locator = cursor{pos: 0}; _ = loc.tell() }",
    nil_composite_promoted_read_call_compile =>
        "package main; type reader interface { read() int }; type loader interface { reader }; func main() { var value loader; _ = value.read() }",
    nil_composite_in_conditional_else_compile =>
        "package main; type doer interface { doWork() }; type actor interface { doer }; func main() { var value actor; if value != nil { value.doWork() } else { _ = 0 } }",
    embedded_interface_promoted_in_return_compile =>
        "package main; type maker interface { make() int }; type builder interface { maker }; type tool struct{}; func (tool) make() int { return 1 }; func build() builder { return tool{} }; func main() { _ = build().make() }",
    overlapping_with_pointer_receiver_compile =>
        "package main; type flushA interface { flush() }; type flushB interface { flush() }; type sink interface { flushA; flushB }; type pipe struct{}; func (p *pipe) flush() {}; func main() { var s sink = &pipe{}; s.flush() }",
    nil_named_interface_method_call_compile =>
        "package main; type handler interface { handle() }; type processor interface { handler }; func invoke(value processor) { value.handle() }; func main() { invoke(nil) }",
    dual_embedded_promoted_call_sites_compile =>
        "package main; type fetch interface { fetch() int }; type store interface { store(int) }; type cache interface { fetch; store }; type mem struct{}; func (mem) fetch() int { return 0 }; func (mem) store(int) {}; func main() { var c cache = mem{}; c.fetch(); c.store(1) }",
}
