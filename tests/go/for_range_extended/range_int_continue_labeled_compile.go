// vybe-test: go/for_range_extended/range_int_continue_labeled_compile
// origin: languages/go/tests/go/test_for_range_extended.rs
// vybe-test-mode: compile

package main
func main() { loop: for i := range 3 { if i == 1 { continue loop }
_ = i } }
