// vybe-test: go/variadic_advanced/variadic_string_spread_compile
// origin: languages/go/tests/go/test_variadic_advanced.rs
// vybe-test-mode: compile

package main
func join(parts ...string) string { s := ""
for _, p := range parts { s += p }
return s }
func main() { tail := []string{"b"}
_ = join(append([]string{"a"}, tail...)...) }
