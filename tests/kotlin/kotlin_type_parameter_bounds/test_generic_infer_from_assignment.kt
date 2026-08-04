// vybe-test: kotlin/kotlin_type_parameter_bounds/test_generic_infer_from_assignment
// origin: languages/kotlin/tests/kotlin/test_kotlin_type_parameter_bounds.rs

class Box<T>(val value: T)

        fun <T> firstOrDefault(value: T?): T {
            return value ?: throw IllegalStateException("missing")
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
            val v: String? = "k"
            val b = Box(firstOrDefault(v))
            __p((b.value).toString())
        
__check("k")
}
