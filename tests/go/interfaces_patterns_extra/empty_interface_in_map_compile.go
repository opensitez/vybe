// vybe-test: go/interfaces_patterns_extra/empty_interface_in_map_compile
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]interface{}{"x": 1}
_ = values }
