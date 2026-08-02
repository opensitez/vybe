// vybe-test: go/struct_tags_reflect/reflect_field_name_reads_tagged_member_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
import "reflect"
type Item struct { SKU string `json:"sku"` }
func main() { _ = reflect.TypeOf(Item{}).Field(0).Name }
