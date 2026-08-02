// vybe-test: kotlin/kotlin_result_runtime/test_result_map_only_on_success
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val mapped = runCatching { 3 }
                .map { it * 10 }
                .map { it + 1 }
            val failed = runCatching<Int> { throw RuntimeException("fail") }
                .map { it + 1 }
            __check((mapped.getOrNull()).toString(), "31")
            __check((failed.isFailure).toString(), "true")
            __check((failed.getOrNull() == null).toString(), "true")
        }
