// vybe-test: kotlin/kotlin_result_runtime/test_result_map_catching_can_transform_failures
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_runtime.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = runCatching { "x".toInt() }
                .mapCatching { it + 1 }
            __check((result.isFailure).toString(), "true")
            val mapped = runCatching { 9 }
                .mapCatching { if (it % 2 == 1) throw IllegalArgumentException("odd") else it }
            __check((mapped.isFailure).toString(), "true")
            __check((mapped.exceptionOrNull()?.message).toString(), "odd")
        }
