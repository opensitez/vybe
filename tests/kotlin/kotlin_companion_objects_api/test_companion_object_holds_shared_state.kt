// vybe-test: kotlin/kotlin_companion_objects_api/test_companion_object_holds_shared_state
// origin: languages/kotlin/tests/kotlin/test_kotlin_companion_objects_api.rs

class Sequence {
            private val id: Int

            private constructor(value: Int) {
                id = value
            }

            companion object {
                var next: Int = 0
                fun nextSequence(): Sequence {
                    next += 1
                    return Sequence(next)
                }
            }

            fun value(): Int = id
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
            val a = Sequence.nextSequence()
            val b = Sequence.nextSequence()
            __p((a.value()).toString())
            __p((b.value()).toString())
            __p((Sequence.next).toString())
        
__check("1\n2\n2")
}
