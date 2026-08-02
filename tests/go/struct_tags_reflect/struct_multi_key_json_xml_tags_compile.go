// vybe-test: go/struct_tags_reflect/struct_multi_key_json_xml_tags_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type Record struct { Title string `json:"title" xml:"title"` }
func main() { _ = Record{} }
