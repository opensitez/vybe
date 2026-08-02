// vybe-test: kotlin/kotlin_sequences_generate/test_generate_sequence_with_null_stop
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

var i = 0
        fun step(value: Int): Int? {
            return if (value < 4) value + 1 else null
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = generateSequence(0) { step(it) }
            __check((seq.toList().joinToString(",")).toString(), "1,2,3,4")
        }
