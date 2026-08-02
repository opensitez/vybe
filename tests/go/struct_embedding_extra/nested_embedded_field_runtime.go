// vybe-test: go/struct_embedding_extra/nested_embedded_field_runtime
// origin: languages/go/tests/go/test_struct_embedding_extra.rs

package main
import "fmt"
type flags struct { enabled bool }
type config struct { flags }
type app struct { config }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := app{config: config{flags: flags{enabled: true}}}
__check(fmt.Sprint(value.enabled), "true")
}
