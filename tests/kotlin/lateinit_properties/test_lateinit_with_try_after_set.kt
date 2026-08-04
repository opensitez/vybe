// vybe-test: kotlin/lateinit_properties/test_lateinit_with_try_after_set
// origin: languages/kotlin/tests/kotlin/test_lateinit_properties.rs

class Probe {
            lateinit var text: String
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
            val p = Probe()
            p.text = "x"
            val result = try {
                p.text
                "after-set"
            } catch (e: Exception) {
                "bad"
            }
            __p((result).toString())
        
__check("after-set")
}
