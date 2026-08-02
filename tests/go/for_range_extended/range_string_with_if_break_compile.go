// vybe-test: go/for_range_extended/range_string_with_if_break_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { for i, r := range "abc" { if i == 1 { break }
_ = r } }
