// vybe-test: kotlin/secondary_constructors/test_secondary_constructor_constructor_can_throw_and_fail
// origin: languages/kotlin/tests/kotlin/test_secondary_constructors.rs

class Guard {
            val value: Int

            constructor(value: Int) {
                if (value < 0) {
                    throw Exception("invalid")
                }
                this.value = value
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
            try {
                __p((Guard(-2).value).toString())
            } catch (e: Exception) {
                __p(("caught").toString())
            }
        
__check("caught")
}
