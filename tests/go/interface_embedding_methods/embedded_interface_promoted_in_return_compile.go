// vybe-test: go/interface_embedding_methods/embedded_interface_promoted_in_return_compile
// origin: languages/go/tests/go/test_interface_embedding_methods.rs
// vybe-test-mode: compile

package main
type maker interface { make() int }
type builder interface { maker }
type tool struct{}
func (tool) make() int { return 1 }
func build() builder { return tool{} }
func main() { _ = build().make() }
