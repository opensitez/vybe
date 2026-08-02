// vybe-test: go/nil_zero_semantics_extra/nil_slice_reslice_compile
// origin: languages/go/tests/go/test_nil_zero_semantics_extra.rs
// vybe-test-mode: compile

package main
func main() { var values []int
_ = values[:] }
