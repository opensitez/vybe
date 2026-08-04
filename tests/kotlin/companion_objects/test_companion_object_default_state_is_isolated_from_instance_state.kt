// vybe-test: kotlin/companion_objects/test_companion_object_default_state_is_isolated_from_instance_state
// origin: languages/kotlin/tests/kotlin/test_companion_objects.rs

class Holder {
            companion object {
                var global = 0
            }

            var local = 0

            init {
                local += 1
                global += local
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
            val first = Holder()
            val second = Holder()
            __p((first.local).toString())
            __p((second.local).toString())
            __p((Holder.global).toString())
        
__check("1\n1\n2")
}
