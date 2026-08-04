// vybe-test: kotlin/generics/test_generic_receiver_extension_with_secondary_projection
// origin: languages/kotlin/tests/kotlin/test_generics.rs

class Holder<T>(private val value: T) {
            fun value(): T = value
        }

        fun <T> Holder<T>.bind(other: T): String {
            return this.value().toString() + ":" + other.toString()
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
            val text = Holder("a")
            val numbers = Holder(4)
            __p((text.bind("x")).toString())
            __p((numbers.bind(6)).toString())
        
__check("a:x\n4:6")
}
