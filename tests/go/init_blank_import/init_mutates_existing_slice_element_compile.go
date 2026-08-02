// vybe-test: go/init_blank_import/init_mutates_existing_slice_element_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
var nums = []int{1, 2, 3}
func init() { nums[1] = 20 }
func main() { _ = nums[1] }
