// vybe-test: kotlin/kotlin_sequences_generate/test_infinite_sequence_is_not_eager
// origin: languages/kotlin/tests/kotlin/test_kotlin_sequences_generate.rs

var calls = 0
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = generateSequence(1) { calls++
it + 1 }
            val taken = values.take(4).toList()
            __check((taken.joinToString(",")).toString(), "1,2,3,4")
            __check((calls >= 1).toString(), "true")
        }
