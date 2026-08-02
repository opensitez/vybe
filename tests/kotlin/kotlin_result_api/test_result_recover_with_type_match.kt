// vybe-test: kotlin/kotlin_result_api/test_result_recover_with_type_match
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching<Int> { throw IllegalArgumentException("x") }
                .recover { cause -> if (cause is IllegalArgumentException) 10 else 0 }
            __check((value.getOrNull()).toString(), "10")
        }
