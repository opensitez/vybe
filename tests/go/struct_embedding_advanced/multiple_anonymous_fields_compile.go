// vybe-test: go/struct_embedding_advanced/multiple_anonymous_fields_compile
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs
// vybe-test-mode: compile

package main
type axis struct { x int }
type ord struct { y int }
type point struct { axis
ord }
func main() { var value point
_ = value.x
_ = value.y }
