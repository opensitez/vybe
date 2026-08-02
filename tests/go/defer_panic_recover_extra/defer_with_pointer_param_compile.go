// vybe-test: go/defer_panic_recover_extra/defer_with_pointer_param_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { value := 1
defer func(ptr *int) { *ptr = 2 }(&value) }
