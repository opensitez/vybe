// vybe-test: go/methods_receivers_extra/method_value_assignment_compile
// origin: languages/go/tests/go/test_methods_receivers_extra.rs
// vybe-test-mode: compile

package main
type counter struct { n int }
func (c counter) total() int { return c.n }
func main() { value := counter{}
fn := value.total
_ = fn }
