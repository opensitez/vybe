// vybe-test: go/range_iteration_extra/range_over_string_blank_identifier_compile
// origin: languages/go/tests/go/test_range_iteration_extra.rs
// vybe-test-mode: compile

package main
func main() { for _, value := range "go" { _ = value } }
