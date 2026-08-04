// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_private_validation
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class PositiveCounter {
            val value: Int

            private constructor(value: Int) {
                this.value = value
            }

            constructor(raw: Int, valid: Boolean) : this(if (valid && raw > 0) raw else 0) {
                __p(("built").toString())
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
            __p((PositiveCounter(-3, true).value).toString())
            __p((PositiveCounter(5, false).value).toString())
            __p((PositiveCounter(5, true).value).toString())
        
__check("built\nbuilt\nbuilt\n0\n0\n5")
}
