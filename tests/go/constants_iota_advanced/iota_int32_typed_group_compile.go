// vybe-test: go/constants_iota_advanced/iota_int32_typed_group_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( I int32 = iota; J; K )
func main() { _, _, _ = I, J, K }
