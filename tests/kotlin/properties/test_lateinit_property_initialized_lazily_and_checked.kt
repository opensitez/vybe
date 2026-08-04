// vybe-test: kotlin/properties/test_lateinit_property_initialized_lazily_and_checked
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Holder {
            lateinit var label: String

            fun initialize(value: String) {
                label = value
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
            val holder = Holder()
            try {
                __p((holder.label).toString())
            } catch (error: UninitializedPropertyAccessException) {
                __p(("not_ready").toString())
            }
            holder.initialize("ok")
            __p((holder.label).toString())
        
__check("not_ready\nok")
}
