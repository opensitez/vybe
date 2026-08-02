// vybe-test: kotlin/kotlin_result_api/test_result_recover_with_unrelated_type
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching<Int> { throw IllegalArgumentException("x") }
                .recover { cause -> if (cause is UnsupportedOperationException) 1 else 2 }
            __check((value.getOrElse { 0 }).toString(), "2")
        }
