// vybe-test: go/constants_iota_advanced/iota_in_array_length_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const size = iota + 3
const ( A = iota; B )
func main() { arr := [size]int{}
_ = arr
_ = A }
