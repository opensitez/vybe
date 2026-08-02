// vybe-test: go/struct_tags_reflect/struct_json_string_option_tag_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type Metric struct { Total int `json:",string"` }
func main() { _ = Metric{} }
