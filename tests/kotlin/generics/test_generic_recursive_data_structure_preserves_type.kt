// vybe-test: kotlin/generics/test_generic_recursive_data_structure_preserves_type
// origin: languages/kotlin/tests/kotlin/test_generics.rs

data class Node<T>(val value: T, val next: Node<T>? = null)

        fun <T> collect(values: Node<T>): String {
            var cursor: Node<T>? = values
            var out = ""
            while (cursor != null) {
                out += cursor.value.toString()
                cursor = cursor.next
                if (cursor != null) out += "-"
            }
            return out
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
            val chain = Node("a", Node("b", Node("c")))
            __p((collect(chain)).toString())
        
__check("a-b-c")
}
