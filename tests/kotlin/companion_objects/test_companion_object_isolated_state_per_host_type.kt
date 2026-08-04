// vybe-test: kotlin/companion_objects/test_companion_object_isolated_state_per_host_type
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Left {
            companion object {
                var value = 1
            }
        }

        class Right {
            companion object {
                var value = 10
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
            Left.value += 1
            Right.value += 5
            __p((Left.value).toString())
            __p((Right.value).toString())
        
__check("2\n15")
}
