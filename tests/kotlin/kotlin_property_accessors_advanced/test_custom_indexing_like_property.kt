// vybe-test: kotlin/kotlin_property_accessors_advanced/test_custom_indexing_like_property
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class Bag {
            private val data = listOf(1, 3, 5)
            operator fun get(index: Int): Int = data[index]
            val head get() = data.first()
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
            val b = Bag()
            __p((b[2]).toString())
            __p((b.head).toString())
        
__check("5\n1")
}
