//! Method sets: value vs pointer receivers, interface satisfaction, embedded promotion.
//! Distinct from `test_method_values.rs` (method values/expressions) and
//! `test_pointer_receivers_advanced.rs` (nil/new/address patterns).

use crate::helpers::*;

go_run_cases! {
    value_receiver_call_on_value_runtime =>
        ("package main; import \"fmt\"; type score struct { pts int }; func (s score) total() int { return s.pts }; func main() { v := score{pts: 11}; fmt.Println(v.total()) }", vec!["11"]),
    value_receiver_call_on_pointer_runtime =>
        ("package main; import \"fmt\"; type score struct { pts int }; func (s score) total() int { return s.pts }; func main() { p := &score{pts: 11}; fmt.Println(p.total()) }", vec!["11"]),
    pointer_receiver_mutates_through_value_variable_runtime =>
        ("package main; import \"fmt\"; type cell struct { n int }; func (c *cell) inc() { c.n++ }; func main() { v := cell{n: 2}; v.inc(); fmt.Println(v.n) }", vec!["3"]),
    pointer_receiver_on_explicit_pointer_runtime =>
        ("package main; import \"fmt\"; type cell struct { n int }; func (c *cell) inc() { c.n++ }; func main() { p := &cell{n: 2}; p.inc(); fmt.Println(p.n) }", vec!["3"]),
    value_receiver_does_not_mutate_field_runtime =>
        ("package main; import \"fmt\"; type cell struct { n int }; func (c cell) bump() { c.n++ }; func main() { v := cell{n: 5}; v.bump(); fmt.Println(v.n) }", vec!["5"]),
    value_type_satisfies_interface_with_value_method_runtime =>
        ("package main; import \"fmt\"; type reader interface { read() int }; type book struct { pages int }; func (b book) read() int { return b.pages }; func main() { var r reader = book{pages: 120}; fmt.Println(r.read()) }", vec!["120"]),
    pointer_type_satisfies_interface_with_value_method_runtime =>
        ("package main; import \"fmt\"; type reader interface { read() int }; type book struct { pages int }; func (b book) read() int { return b.pages }; func main() { var r reader = &book{pages: 120}; fmt.Println(r.read()) }", vec!["120"]),
    pointer_type_satisfies_interface_with_pointer_method_runtime =>
        ("package main; import \"fmt\"; type writer interface { write(int) }; type pad struct { n int }; func (p *pad) write(v int) { p.n = v }; func main() { var w writer = &pad{}; w.write(9); fmt.Println(w.(*pad).n) }", vec!["9"]),
    value_type_does_not_implement_pointer_only_interface_runtime =>
        ("package main; import \"fmt\"; type mutator interface { set(int) }; type gauge struct { n int }; func (g *gauge) set(v int) { g.n = v }; func main() { g := gauge{}; var m mutator = &g; m.set(4); fmt.Println(g.n) }", vec!["4"]),
    embedded_value_method_promoted_on_outer_runtime =>
        ("package main; import \"fmt\"; type base struct { id int }; func (b base) idVal() int { return b.id }; type shell struct { base }; func main() { s := shell{base: base{id: 7}}; fmt.Println(s.idVal()) }", vec!["7"]),
    embedded_pointer_method_promoted_on_outer_runtime =>
        ("package main; import \"fmt\"; type base struct { n int }; func (b *base) double() { b.n *= 2 }; type shell struct { base }; func main() { s := shell{base: base{n: 3}}; s.double(); fmt.Println(s.n) }", vec!["6"]),
    embedded_pointer_field_method_promoted_runtime =>
        ("package main; import \"fmt\"; type inner struct { tag string }; func (i inner) label() string { return i.tag }; type outer struct { *inner }; func main() { o := outer{inner: &inner{tag: \"go\"}}; fmt.Println(o.label()) }", vec!["go"]),
    dual_receivers_value_read_pointer_write_runtime =>
        ("package main; import \"fmt\"; type wallet struct { cash int }; func (w wallet) balance() int { return w.cash }; func (w *wallet) deposit(v int) { w.cash += v }; func main() { w := wallet{cash: 10}; fmt.Println(w.balance()); w.deposit(5); fmt.Println(w.balance()) }", vec!["10", "15"]),
    method_on_defined_int_type_value_receiver_runtime =>
        ("package main; import \"fmt\"; type counter int; func (c counter) next() int { return int(c) + 1 }; func main() { var c counter = 4; fmt.Println(c.next()) }", vec!["5"]),
    method_on_defined_int_type_pointer_receiver_runtime =>
        ("package main; import \"fmt\"; type counter int; func (c *counter) add(v int) { *c += counter(v) }; func main() { var c counter = 4; c.add(3); fmt.Println(int(c)) }", vec!["7"]),
    interface_assign_from_address_of_local_runtime =>
        ("package main; import \"fmt\"; type saver interface { save(int) }; type disk struct { used int }; func (d *disk) save(v int) { d.used += v }; func main() { local := disk{}; var s saver = &local; s.save(6); fmt.Println(local.used) }", vec!["6"]),
    value_method_on_nonaddressable_temp_via_interface_runtime =>
        ("package main; import \"fmt\"; type speaker interface { say() string }; type bot struct{}; func (b bot) say() string { return \"beep\" }; func main() { var s speaker = bot{}; fmt.Println(s.say()) }", vec!["beep"]),
    promoted_embedded_overrides_outer_field_access_runtime =>
        ("package main; import \"fmt\"; type inner struct { x int }; type outer struct { inner; x int }; func main() { o := outer{inner: inner{x: 1}, x: 2}; fmt.Println(o.x); fmt.Println(o.inner.x) }", vec!["2", "1"]),
    nested_embedded_value_method_two_levels_runtime =>
        ("package main; import \"fmt\"; type leaf struct { v int }; func (l leaf) val() int { return l.v }; type branch struct { leaf }; type trunk struct { branch }; func main() { t := trunk{branch: branch{leaf: leaf{v: 9}}}; fmt.Println(t.val()) }", vec!["9"]),
    pointer_embedded_nil_inner_method_call_panic_guard_runtime =>
        ("package main; import \"fmt\"; type inner struct { n int }; func (i *inner) peek() int { if i == nil { return -1 }; return i.n }; type outer struct { *inner }; func main() { var o outer; fmt.Println(o.peek()) }", vec!["-1"]),
    interface_boxing_value_then_pointer_method_set_runtime =>
        ("package main; import \"fmt\"; type grower interface { grow() }; type plant struct { h int }; func (p *plant) grow() { p.h++ }; func main() { p := &plant{h: 1}; var g grower = p; g.grow(); fmt.Println(p.h) }", vec!["2"]),
    value_receiver_string_method_runtime =>
        ("package main; import \"fmt\"; type label struct { text string }; func (l label) upper() string { return l.text + \"!\" }; func main() { fmt.Println(label{text: \"hi\"}.upper()) }", vec!["hi!"]),
    pointer_receiver_slice_header_mutation_runtime =>
        ("package main; import \"fmt\"; type bag struct { items []int }; func (b *bag) appendItem(v int) { b.items = append(b.items, v) }; func main() { b := bag{items: []int{1}}; b.appendItem(2); fmt.Println(len(b.items)); fmt.Println(b.items[1]) }", vec!["2", "2"]),
    value_receiver_slice_header_no_mutation_runtime =>
        ("package main; import \"fmt\"; type bag struct { items []int }; func (b bag) appendItem(v int) { b.items = append(b.items, v) }; func main() { b := bag{items: []int{1}}; b.appendItem(2); fmt.Println(len(b.items)) }", vec!["1"]),
    embedded_anonymous_struct_method_promotion_runtime =>
        ("package main; import \"fmt\"; type coords struct { x int; y int }; func (c coords) sum() int { return c.x + c.y }; type point struct { coords }; func main() { p := point{coords: coords{x: 2, y: 5}}; fmt.Println(p.sum()) }", vec!["7"]),
    method_expression_value_receiver_from_type_runtime =>
        ("package main; import \"fmt\"; type pair struct { a int }; func (p pair) first() int { return p.a }; func main() { fn := pair.first; fmt.Println(fn(pair{a: 8})) }", vec!["8"]),
    method_expression_pointer_receiver_from_type_runtime =>
        ("package main; import \"fmt\"; type pair struct { a int }; func (p *pair) set(v int) { p.a = v }; func main() { target := &pair{}; fn := (*pair).set; fn(target, 6); fmt.Println(target.a) }", vec!["6"]),
    interface_satisfied_by_embedded_promoted_method_runtime =>
        ("package main; import \"fmt\"; type runner interface { run() int }; type legs struct{}; func (legs) run() int { return 42 }; type athlete struct { legs }; func main() { var r runner = athlete{}; fmt.Println(r.run()) }", vec!["42"]),
    pointer_to_value_still_gets_value_method_set_runtime =>
        ("package main; import \"fmt\"; type tile struct { color string }; func (t tile) hue() string { return t.color }; func main() { t := tile{color: \"red\"}; p := &t; fmt.Println(p.hue()) }", vec!["red"]),
    value_with_only_pointer_methods_needs_address_runtime =>
        ("package main; import \"fmt\"; type latch struct { on bool }; func (l *latch) flip() { l.on = !l.on }; func main() { l := latch{on: false}; l.flip(); fmt.Println(l.on) }", vec!["true"]),
    embedded_value_type_pointer_method_on_outer_value_runtime =>
        ("package main; import \"fmt\"; type engine struct { rpm int }; func (e *engine) rev() { e.rpm++ }; type car struct { engine }; func main() { c := car{engine: engine{rpm: 1000}}; c.rev(); fmt.Println(c.rpm) }", vec!["1001"]),
    interface_from_literal_pointer_with_pointer_method_runtime =>
        ("package main; import \"fmt\"; type resetter interface { reset() }; type timer struct { ticks int }; func (t *timer) reset() { t.ticks = 0 }; func main() { var r resetter = &timer{ticks: 5}; r.reset(); fmt.Println(r.(*timer).ticks) }", vec!["0"]),
    multiple_embedded_types_distinct_promoted_methods_runtime =>
        ("package main; import \"fmt\"; type north struct{}; func (north) dir() string { return \"N\" }; type east struct{}; func (east) dir() string { return \"E\" }; type compass struct { north; east }; func main() { c := compass{}; fmt.Println(c.north.dir()); fmt.Println(c.east.dir()) }", vec!["N", "E"]),
    value_receiver_bool_toggle_copy_runtime =>
        ("package main; import \"fmt\"; type flag struct { on bool }; func (f flag) isOn() bool { return f.on }; func main() { f := flag{on: true}; fmt.Println(f.isOn()) }", vec!["true"]),
    pointer_receiver_chain_returns_pointer_runtime =>
        ("package main; import \"fmt\"; type node struct { val int }; func (n *node) add(v int) *node { n.val += v; return n }; func main() { n := &node{val: 1}; n.add(2).add(3); fmt.Println(n.val) }", vec!["6"]),
    embedded_interface_field_method_dispatch_runtime =>
        ("package main; import \"fmt\"; type speaker interface { talk() string }; type bot struct{}; func (bot) talk() string { return \"bot\" }; type host struct { speaker }; func main() { h := host{speaker: bot{}}; fmt.Println(h.talk()) }", vec!["bot"]),
    defined_type_underlying_struct_value_method_runtime =>
        ("package main; import \"fmt\"; type meters float64; func (m meters) km() float64 { return float64(m) / 1000 }; func main() { fmt.Println(meters(2500).km()) }", vec!["2.5"]),
    defined_type_underlying_struct_pointer_method_runtime =>
        ("package main; import \"fmt\"; type meters float64; func (m *meters) scale(f float64) { *m = meters(float64(*m) * f) }; func main() { var m meters = 100; m.scale(2); fmt.Println(float64(m)) }", vec!["200"]),
    outer_method_shadows_embedded_method_set_runtime =>
        ("package main; import \"fmt\"; type base struct{}; func (base) tag() string { return \"base\" }; type derived struct { base }; func (derived) tag() string { return \"derived\" }; func main() { d := derived{}; fmt.Println(d.tag()) }", vec!["derived"]),
    explicit_embedded_type_qualifier_method_call_runtime =>
        ("package main; import \"fmt\"; type base struct{}; func (base) tag() string { return \"base\" }; type derived struct { base }; func (derived) tag() string { return \"derived\" }; func main() { d := derived{}; fmt.Println(d.base.tag()) }", vec!["base"]),
    interface_value_type_with_both_method_sets_runtime =>
        ("package main; import \"fmt\"; type speaker interface { say() string }; type bot struct { msg string }; func (b bot) say() string { return b.msg }; func main() { var s speaker = bot{msg: \"hi\"}; fmt.Println(s.say()) }", vec!["hi"]),
    pointer_only_interface_from_local_address_runtime =>
        ("package main; import \"fmt\"; type mutator interface { mutate() }; type data struct { n int }; func (d *data) mutate() { d.n = 99 }; func main() { local := data{n: 1}; var m mutator = &local; m.mutate(); fmt.Println(local.n) }", vec!["99"]),
}

go_compile_cases! {
    value_type_assign_to_value_method_interface_compile =>
        "package main; type worker interface { work() int }; type task struct { id int }; func (t task) work() int { return t.id }; func main() { var w worker = task{id: 1}; _ = w }",
    pointer_type_assign_to_value_method_interface_compile =>
        "package main; type worker interface { work() int }; type task struct { id int }; func (t task) work() int { return t.id }; func main() { var w worker = &task{id: 1}; _ = w }",
    pointer_only_method_interface_needs_pointer_compile =>
        "package main; type editor interface { edit() }; type doc struct{}; func (d *doc) edit() {}; func main() { var e editor = &doc{}; _ = e }",
    embedded_promoted_method_call_compile =>
        "package main; type inner struct{}; func (inner) f() {}; type outer struct { inner }; func main() { var o outer; o.f() }",
    embedded_pointer_field_promoted_method_compile =>
        "package main; type inner struct{}; func (inner) f() {}; type outer struct { *inner }; func main() { o := outer{inner: &inner{}}; o.f() }",
    method_set_value_expression_compile =>
        "package main; type T struct{}; func (T) M() {}; func main() { _ = T.M }",
    method_set_pointer_expression_compile =>
        "package main; type T struct{}; func (t *T) M() {}; func main() { _ = (*T).M }",
    address_of_temp_for_pointer_method_compile =>
        "package main; type cell struct { n int }; func (c *cell) set(v int) { c.n = v }; func main() { c := cell{}; c.set(1) }",
    two_level_embedded_method_promotion_compile =>
        "package main; type leaf struct{}; func (leaf) f() {}; type branch struct { leaf }; type trunk struct { branch }; func main() { trunk{}.f() }",
    interface_field_with_embedded_implementer_compile =>
        "package main; type doer interface { doWork() }; type impl struct{}; func (impl) doWork() {}; type holder struct { doer }; func main() { _ = holder{doer: impl{}} }",
    value_receiver_on_map_value_type_compile =>
        "package main; type key struct { s string }; func (k key) hash() int { return len(k.s) }; func main() { _ = key{s: \"a\"}.hash() }",
    pointer_receiver_on_slice_element_address_compile =>
        "package main; type item struct { n int }; func (i *item) bump() { i.n++ }; func main() { items := []item{{n: 1}}; items[0].bump() }",
    dual_embedded_same_method_name_requires_qualifier_compile =>
        "package main; type a struct{}; func (a) f() {}; type b struct{}; func (b) f() {}; type c struct { a; b }; func main() { var x c; x.a.f(); x.b.f() }",
    named_interface_from_embedded_method_set_compile =>
        "package main; type mover interface { move() int }; type legs struct{}; func (legs) move() int { return 1 }; type body struct { legs }; func main() { var m mover = body{}; _ = m }",
    pointer_embedded_nil_safe_explicit_access_compile =>
        "package main; type inner struct { n int }; type outer struct { *inner }; func main() { var o outer; _ = o.inner }",
}

macro_rules! go_compile_fail_cases {
    ($($name:ident => $src:expr,)+) => {
        $(#[test] fn $name() { assert!(!compile_ok_check($src)); })+
    };
}

go_compile_fail_cases! {
    value_type_cannot_assign_pointer_only_interface_compile_fail =>
        "package main; type editor interface { edit() }; type doc struct{}; func (d *doc) edit() {}; func main() { var e editor = doc{}; _ = e }",
    value_variable_pointer_only_method_expression_compile_fail =>
        "package main; type cell struct { n int }; func (c *cell) set(v int) { c.n = v }; func main() { v := cell{}; fn := cell.set; _ = fn }",
    ambiguous_promoted_method_without_qualifier_compile_fail =>
        "package main; type a struct{}; func (a) f() {}; type b struct{}; func (b) f() {}; type c struct { a; b }; func main() { var x c; x.f() }",
}
