// vybe-test: go/blank_identifier_extended/blank_interface_embedded_method_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
type Reader interface { Read(p []byte) (n int, err error) }
type rw struct{}
func (r rw) Read(p []byte) (int, error) { return 0, nil }
func main() { var rd Reader = rw{}
_ = rd }
