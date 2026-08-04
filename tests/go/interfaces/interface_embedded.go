// vybe-test: go/interfaces/interface_embedded
// origin: languages/go/tests/go/test_interfaces.rs
// vybe-test-mode: compile

package main
type Reader interface { Read() string } type Writer interface { Write(s string) } type ReadWriter interface { Reader
Writer } func main() {}
