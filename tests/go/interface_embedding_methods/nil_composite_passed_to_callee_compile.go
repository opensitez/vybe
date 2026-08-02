// vybe-test: go/interface_embedding_methods/nil_composite_passed_to_callee_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type speaker interface { speak() string }
type talker interface { speaker }
func say(value talker) string { if value == nil { return "nil" }
return value.speak() }
func main() { _ = say(nil) }
