// vybe-test: go/structs/struct_literal_partial
// origin: languages/go/tests/go/test_structs.rs
// vybe-test-mode: compile

package main
type Config struct { Host string
Port int
Debug bool } func main() { c := Config{Host: "localhost"}
_ = c }
