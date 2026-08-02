// vybe-test: go/variadic_advanced/variadic_type_switch_on_interface_pack_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func classify(values ...interface{}) int { c := 0
for _, v := range values { switch v.(type) { case int: c++ } }
return c }
func main() { _ = classify(1, "x", 2) }
