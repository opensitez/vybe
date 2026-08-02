// vybe-test: go/constants_iota_advanced/iota_typed_int64_group_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( N0 int64 = iota; N1; N2 )
func main() { _, _, _ = N0, N1, N2 }
