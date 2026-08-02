// vybe-test: go/init_blank_import/init_chain_reads_prior_init_value_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
var first int
var second int
func init() { first = 3 }
func init() { second = first + 2 }
func main() { _ = second }
