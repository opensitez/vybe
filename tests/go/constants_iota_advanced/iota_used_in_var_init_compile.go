// vybe-test: go/constants_iota_advanced/iota_used_in_var_init_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( A = iota; B )
var total = A + B
func main() { _ = total }
