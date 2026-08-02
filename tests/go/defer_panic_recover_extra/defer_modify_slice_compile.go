// vybe-test: go/defer_panic_recover_extra/defer_modify_slice_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { values := []int{1}
defer func() { values[0] = 2 }() }
