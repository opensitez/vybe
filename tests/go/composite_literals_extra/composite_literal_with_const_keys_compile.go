// vybe-test: go/composite_literals_extra/composite_literal_with_const_keys_compile
// origin: languages/go/tests/go/test_composite_literals_extra.rs
// vybe-test-mode: compile

package main
const home = "home"
func main() { values := map[string]int{home: 1}
_ = values }
