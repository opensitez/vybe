// vybe-test: kotlin/companion_objects/test_companion_object_state_mutation_is_shared_with_factory_calls
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Counter {
            companion object {
                private var next = 0

                fun next(delta: Int = 1): Int {
                    next += delta
                    return next
                }

                fun current(): Int = next
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
            __p((Counter.next()).toString())
            __p((Counter.next(3)).toString())
            __p((Counter.next()).toString())
            __p((Counter.current()).toString())
        
__check("1\n4\n5\n5")
}
