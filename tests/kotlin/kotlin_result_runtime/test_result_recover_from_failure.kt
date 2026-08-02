// vybe-test: kotlin/kotlin_result_runtime/test_result_recover_from_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = runCatching { throw IllegalArgumentException("bad") }
                .recover { 12 }
            val failed = runCatching<Int> { throw IllegalStateException("boom") }
                .recover { cause -> if (cause is IllegalArgumentException) -1 else 99 }
            __check((result.getOrNull()).toString(), "12")
            __check((failed.getOrNull()).toString(), "99")
        }
