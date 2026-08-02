// vybe-test: go/defer_panic_recover_extra/defer_modify_map_compile
// origin: languages/go/tests/go/test_defer_panic_recover_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{"a": 1}
defer func() { values["a"] = 2 }() }
