// vybe-test: go/struct_tags_reflect/reflect_field_by_name_func_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Row struct { Alpha int `json:"alpha"`
Beta int `json:"beta"` }
func main() { _, _ = reflect.TypeOf(Row{}).FieldByNameFunc(func(name string) bool { return len(name) == 4 }) }
