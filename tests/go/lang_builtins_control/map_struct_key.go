// vybe-test: go/lang_builtins_control/map_struct_key
// origin: languages/go/tests/go/test_lang_builtins_control.rs
// vybe-test-mode: compile

package main
type K struct { X int }
func main() { _ = map[K]string{} }
