// vybe-test: go/nil_zero_semantics_extra/nil_interface_in_map_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]interface{}{"x": nil}
_ = values }
