// vybe-test: kotlin/kotlin_result_runtime/test_result_recover_catching
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = runCatching { throw IllegalArgumentException("bad") }
                .recoverCatching { throw IllegalStateException("wrapped") }
            __check((result.isFailure).toString(), "true")
            __check((result.exceptionOrNull()?.let { it::class.simpleName }).toString(), "IllegalStateException")
        }
