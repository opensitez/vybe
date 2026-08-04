// vybe-test: kotlin/type_aliases/test_typealias_sequence_operations_preserve_lazy_contract
// origin: languages/kotlin/tests/kotlin/test_type_aliases.rs

typealias IntSequence = Sequence<Int>

        fun firstSquares(limit: Int): IntSequence {
            return generateSequence(0) { value ->
                if (value + 2 <= limit) value + 2 else null
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
            val values = firstSquares(6)
            __p((values.take(3).joinToString(",")).toString())
            __p((firstSquares(6).sum()).toString())
        
__check("2,4,6\n12")
}
