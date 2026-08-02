// vybe-test: go/type_aliases/method_on_defined_struct_type
// origin: languages/go/tests/go/test_type_aliases.rs
// vybe-test-mode: compile

package main
type Row struct { n int }
func (r Row) total() int { return r.n }
func main() { _ = Row{n: 1}.total() }
