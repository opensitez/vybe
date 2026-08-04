// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_with_nested_class_argument
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Host {
            val name: String
            class Config

            constructor(value: String) {
                this.name = value
            }

            constructor(config: Config, value: String) : this(value) {
                val used = config
                __p((used is Config).toString())
            }
        }

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __p((Host("root").name).toString())
            __p((Host(Host.Config(), "inner").name).toString())
        
__check("true\nroot\ninner")
}
