// vybe-test: go/interfaces/interface_nil_check
// origin: languages/go/tests/go/test_interfaces.rs
// vybe-test-mode: compile

package main
type Doer interface { Do() }
func run(d Doer) { if d == nil { return }
d.Do() } func main() { run(nil) }
