// vybe-test: go/lang_builtins_control/slice_to_array_ptr
// origin: languages/go/tests/go/test_lang_builtins_control.rs
// vybe-test-mode: compile

package main
func main() { s := []int{1,2}
_ = (*[2]int)(s) }
