// vybe-test: go/type_conversions_extra/conversion_of_map_lookup_compile
// origin: languages/go/tests/go/test_type_conversions_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{"a": 7}
_ = float64(values["a"]) }
