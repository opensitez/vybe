// vybe-test: kotlin/kotlin_result_api/test_result_recover_catching_flips_failure
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bad = runCatching<Int> { throw IllegalArgumentException("bad") }
            val recovered = bad.recoverCatching { throw IllegalStateException("wrapped") }
            __check((recovered.isFailure).toString(), "true")
            __check((recovered.exceptionOrNull()?.let { it::class.simpleName }).toString(), "IllegalStateException")
        }
