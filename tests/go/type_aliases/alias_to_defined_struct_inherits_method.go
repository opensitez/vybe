// vybe-test: go/type_aliases/alias_to_defined_struct_inherits_method
// origin: languages/go/tests/go/test_type_aliases.rs
// vybe-test-mode: compile

package main
type Row struct { n int }
func (r Row) total() int { return r.n }
type Alias = Row
func main() { _ = Alias{n: 2}.total() }
