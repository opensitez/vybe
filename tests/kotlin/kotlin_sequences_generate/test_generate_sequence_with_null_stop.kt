// vybe-test: kotlin/kotlin_sequences_generate/test_generate_sequence_with_null_stop
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

var i = 0
        fun step(value: Int): Int? {
            return if (value < 4) value + 1 else null
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
            val seq = generateSequence(0) { step(it) }
            __p((seq.toList().joinToString(",")).toString())
        
__check("1,2,3,4")
}
