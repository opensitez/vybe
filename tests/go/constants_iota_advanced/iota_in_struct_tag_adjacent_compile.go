// vybe-test: go/constants_iota_advanced/iota_in_struct_tag_adjacent_compile
// origin: languages/go/tests/go/test_constants_iota_advanced.rs
// vybe-test-mode: compile

package main
const ( Tag = iota; Other )
type item struct { id int }
func main() { _ = Tag + Other }
