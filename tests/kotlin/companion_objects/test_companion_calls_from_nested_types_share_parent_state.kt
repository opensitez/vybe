// vybe-test: kotlin/companion_objects/test_companion_calls_from_nested_types_share_parent_state
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Board {
            companion object {
                private var tokens = 0
                fun hit(): Int {
                    tokens += 1
                    return tokens
                }
            }

            class Checker {
                fun hit(): Int = Board.hit()
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
            __p((Board.hit()).toString())
            __p((Board.Checker().hit()).toString())
            __p((Board.hit()).toString())
        
__check("1\n2\n3")
}
