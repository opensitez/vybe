use crate::helpers::*;

macro_rules! go_compile_test {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

macro_rules! go_run_test {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let out = run_prints($src);
            assert_eq!(out, $expected);
        }
    };
}

macro_rules! run_cases {
    ($( $name:ident => ($src:expr, $expected:expr), )*) => {
        $( go_run_test!($name, $src, $expected); )*
    };
}

macro_rules! compile_cases {
    ($( $name:ident => $src:expr, )*) => {
        $( go_compile_test!($name, $src); )*
    };
}

run_cases! {
    dual_receiver_value_read_pointer_write_runtime =>
        ("package main; import \"fmt\"; type ledger struct { balance int }; func (l ledger) snapshot() int { return l.balance }; func (l *ledger) deposit(v int) { l.balance += v }; func main() { value := ledger{balance: 10}; fmt.Println(value.snapshot()); value.deposit(5); fmt.Println(value.snapshot()); }", vec!["10", "15"]),

    value_receiver_struct_field_copy_not_mutated_runtime =>
        ("package main; import \"fmt\"; type cell struct { n int }; func (c cell) bump() { c.n++ }; func main() { value := cell{n: 4}; value.bump(); fmt.Println(value.n); }", vec!["4"]),

    pointer_receiver_struct_field_mutated_runtime =>
        ("package main; import \"fmt\"; type cell struct { n int }; func (c *cell) bump() { c.n++ }; func main() { value := cell{n: 4}; value.bump(); fmt.Println(value.n); }", vec!["5"]),

    value_receiver_slice_append_not_visible_runtime =>
        ("package main; import \"fmt\"; type bag struct { items []int }; func (b bag) push(v int) { b.items = append(b.items, v) }; func main() { value := bag{items: []int{1}}; value.push(2); fmt.Println(len(value.items)); }", vec!["1"]),

    pointer_receiver_slice_append_visible_runtime =>
        ("package main; import \"fmt\"; type bag struct { items []int }; func (b *bag) push(v int) { b.items = append(b.items, v) }; func main() { value := bag{items: []int{1}}; value.push(2); fmt.Println(len(value.items)); fmt.Println(value.items[1]); }", vec!["2", "2"]),

    nil_pointer_receiver_nil_check_without_panic_runtime =>
        ("package main; import \"fmt\"; type node struct { id int }; func (n *node) absent() bool { return n == nil }; func main() { var value *node; fmt.Println(value.absent()); }", vec!["true"]),

    address_of_field_then_pointer_method_runtime =>
        ("package main; import \"fmt\"; type meter struct { reading int }; func (m *meter) set(v int) { m.reading = v }; func main() { value := meter{reading: 0}; fieldPtr := &value.reading; *fieldPtr = 3; value.set(7); fmt.Println(value.reading); }", vec!["7"]),

    new_struct_zero_value_matches_literal_runtime =>
        ("package main; import \"fmt\"; type widget struct { size int; label string }; func main() { fromNew := new(widget); fromLit := &widget{}; fmt.Println(fromNew.size == fromLit.size); fmt.Println(fromNew.label == fromLit.label); }", vec!["true", "true"]),

    new_struct_pointer_receiver_zero_init_runtime =>
        ("package main; import \"fmt\"; type tally struct { sum int }; func (t *tally) add(v int) { t.sum += v }; func main() { value := new(tally); value.add(3); value.add(4); fmt.Println(value.sum); }", vec!["7"]),

    literal_pointer_struct_pointer_receiver_init_runtime =>
        ("package main; import \"fmt\"; type tally struct { sum int }; func (t *tally) add(v int) { t.sum += v }; func main() { value := &tally{sum: 1}; value.add(4); fmt.Println(value.sum); }", vec!["5"]),

    new_vs_literal_same_method_result_runtime =>
        ("package main; import \"fmt\"; type score struct { points int }; func (s *score) double() { s.points = s.points * 2 }; func main() { a := new(score); a.points = 5; a.double(); b := &score{points: 5}; b.double(); fmt.Println(a.points); fmt.Println(b.points); }", vec!["10", "10"]),

    new_assign_fields_equals_keyed_literal_runtime =>
        ("package main; import \"fmt\"; type point struct { x int; y int }; func main() { fromNew := new(point); fromNew.x = 3; fromNew.y = 8; fromLit := &point{x: 3, y: 8}; fmt.Println(fromNew.x == fromLit.x); fmt.Println(fromNew.y == fromLit.y); }", vec!["true", "true"]),

    new_pointer_receiver_chain_mutation_runtime =>
        ("package main; import \"fmt\"; type chain struct { total int }; func (c *chain) step(v int) *chain { c.total += v; return c }; func main() { value := new(chain); value.step(2).step(5); fmt.Println(value.total); }", vec!["7"]),

    literal_pointer_receiver_chain_mutation_runtime =>
        ("package main; import \"fmt\"; type chain struct { total int }; func (c *chain) step(v int) *chain { c.total += v; return c }; func main() { value := &chain{total: 1}; value.step(2).step(5); fmt.Println(value.total); }", vec!["8"]),

    pointer_receiver_via_field_address_runtime =>
        ("package main; import \"fmt\"; type holder struct { gauge int }; func (h *holder) raise() { h.gauge++ }; func main() { value := holder{gauge: 2}; alias := &value; alias.raise(); fmt.Println(value.gauge); }", vec!["3"]),

    value_method_then_pointer_method_sequence_runtime =>
        ("package main; import \"fmt\"; type account struct { balance int }; func (a account) funds() int { return a.balance }; func (a *account) credit(v int) { a.balance += v }; func main() { value := account{balance: 20}; fmt.Println(value.funds()); value.credit(7); fmt.Println(value.funds()); }", vec!["20", "27"]),

    method_expression_value_receiver_on_type_runtime =>
        ("package main; import \"fmt\"; type tag struct { name string }; func (t tag) label() string { return t.name }; func main() { fn := tag.label; fmt.Println(fn(tag{name: \"go\"})); }", vec!["go"]),

    method_expression_pointer_receiver_on_type_runtime =>
        ("package main; import \"fmt\"; type tag struct { name string }; func (t *tag) rename(v string) { t.name = v }; func main() { value := &tag{name: \"old\"}; fn := (*tag).rename; fn(value, \"new\"); fmt.Println(value.name); }", vec!["new"]),

    pointer_receiver_bool_field_toggle_runtime =>
        ("package main; import \"fmt\"; type flag struct { on bool }; func (f *flag) flip() { f.on = !f.on }; func main() { value := &flag{on: false}; value.flip(); fmt.Println(value.on); }", vec!["true"]),

    value_receiver_read_after_pointer_write_runtime =>
        ("package main; import \"fmt\"; type note struct { text string }; func (n *note) set(v string) { n.text = v }; func (n note) read() string { return n.text }; func main() { value := note{text: \"a\"}; value.set(\"b\"); fmt.Println(value.read()); }", vec!["b"]),
}

compile_cases! {
    nil_pointer_receiver_field_deref_compile =>
        "package main; type node struct { id int }; func (n *node) read() int { return n.id }; func main() { var value *node; _ = value.read() }",

    nil_pointer_value_receiver_field_deref_compile =>
        "package main; type node struct { id int }; func (n node) read() int { return n.id }; func main() { var value *node; _ = value.read() }",

    nil_receiver_recover_wrapper_compile =>
        "package main; type widget struct { size int }; func (w *widget) read() int { return w.size }; func safe(v *widget) { defer func() { recover() }(); _ = v.read() }; func main() { var value *widget; safe(value) }",

    address_of_struct_field_assignment_compile =>
        "package main; type pair struct { x int; y int }; func main() { value := pair{x: 1, y: 2}; ptr := &value.y; *ptr = 5; _ = value }",

    address_of_field_in_pointer_receiver_method_compile =>
        "package main; type pair struct { x int; y int }; func (p *pair) swap() { px := &p.x; py := &p.y; *px, *py = *py, *px }; func main() { value := pair{x: 1, y: 2}; value.swap() }",

    new_and_literal_pointer_receiver_compile =>
        "package main; type tally struct { sum int }; func (t *tally) add(v int) { t.sum += v }; func main() { a := new(tally); b := &tally{}; a.add(1); b.add(2) }",

    dual_receiver_method_expression_compile =>
        "package main; type ledger struct { balance int }; func (l ledger) snapshot() int { return l.balance }; func (l *ledger) deposit(v int) { l.balance += v }; func main() { _ = ledger.snapshot; _ = (*ledger).deposit }",

    pointer_receiver_takes_field_address_compile =>
        "package main; type holder struct { gauge int }; func (h *holder) reset() { target := &h.gauge; *target = 0 }; func main() { value := &holder{gauge: 3}; value.reset() }",

    value_receiver_on_named_type_pointer_compile =>
        "package main; type degree int; func (d degree) sign() int { if d < 0 { return -1 }; if d > 0 { return 1 }; return 0 }; func main() { var value *degree; _ = value.sign() }",

    new_pointer_passed_to_value_receiver_compile =>
        "package main; type token struct { id int }; func (t token) value() int { return t.id }; func main() { created := new(token); _ = created.value() }",

    literal_pointer_passed_to_value_receiver_compile =>
        "package main; type token struct { id int }; func (t token) value() int { return t.id }; func main() { created := &token{id: 4}; _ = created.value() }",

    address_of_field_in_composite_literal_variable_compile =>
        "package main; type row struct { cells [2]int }; func main() { value := row{cells: [2]int{1, 2}}; ptr := &value.cells[1]; _ = ptr }",

    nil_pointer_method_value_binding_compile =>
        "package main; type widget struct { size int }; func (w *widget) grow() { w.size++ }; func main() { var value *widget; fn := value.grow; fn() }",

    address_of_int_field_mutate_compile =>
        "package main; type pair struct { x int; y int }; func main() { value := pair{x: 1, y: 2}; ptr := &value.x; *ptr = 9; _ = value }",

    address_of_field_passed_to_mutator_compile =>
        "package main; func scale(target *int, factor int) { *target = *target * factor }; func main() { value := struct{ n int }{n: 3}; scale(&value.n, 4) }",

    address_of_nested_struct_field_compile =>
        "package main; type inner struct { count int }; type outer struct { core inner }; func main() { value := outer{core: inner{count: 6}}; ptr := &value.core.count; *ptr = 11; _ = value }",
}
