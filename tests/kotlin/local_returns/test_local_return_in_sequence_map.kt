// vybe-test: kotlin/local_returns/test_local_return_in_sequence_map
// origin: languages/kotlin/tests/kotlin/test_local_returns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = sequenceOf(1, 2, 3).map {
                if (it == 1) return@map 10
                it
            }.toList().joinToString(",")
            __check((out).toString(), "10,2,3")
        }
