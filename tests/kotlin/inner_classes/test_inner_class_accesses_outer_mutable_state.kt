// vybe-test: kotlin/inner_classes/test_inner_class_accesses_outer_mutable_state
// origin: languages/kotlin/tests/kotlin/test_inner_classes.rs

class Store {
            var value = 1
            inner class Bump {
                fun add(v: Int) {
                    value += v
                }
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
            val store = Store()
            val bump = store.Bump()
            bump.add(5)
            __p((store.value).toString())
        
__check("6")
}
