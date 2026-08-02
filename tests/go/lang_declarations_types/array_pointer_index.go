// vybe-test: go/lang_declarations_types/array_pointer_index
// origin: languages/go/tests/go/test_lang_declarations_types.rs
// vybe-test-mode: compile

package main
func main() { a := [2]int{1,2}
p := &a
_ = (*p)[1] }
