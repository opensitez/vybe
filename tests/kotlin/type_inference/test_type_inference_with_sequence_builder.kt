// vybe-test: kotlin/type_inference/test_type_inference_with_sequence_builder
// origin: languages/kotlin/tests/kotlin/test_type_inference.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seq = generateSequence(1) { if (it < 4) it + 1 else null }
            __check((seq.toList().joinToString(",")).toString(), "1,2,3,4")
        }
