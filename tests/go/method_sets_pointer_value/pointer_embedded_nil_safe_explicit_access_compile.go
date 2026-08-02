// vybe-test: go/method_sets_pointer_value/pointer_embedded_nil_safe_explicit_access_compile
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs
// vybe-test-mode: compile

package main
type inner struct { n int }
type outer struct { *inner }
func main() { var o outer
_ = o.inner }
