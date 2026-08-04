// vybe-test: go/type_assertions/interface_to_interface_assert
// origin: languages/go/tests/go/test_type_assertions.rs
// vybe-test-mode: compile

package main
type Reader interface { Read() }
type Writer interface { Write() }
type ReadWriter interface { Reader
Writer }
func main() { var rw ReadWriter
var r Reader = rw
_ = r.(Writer)
}
