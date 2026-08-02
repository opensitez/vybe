// vybe-test: kotlin/result_patterns/test_result_map_catching_catches_mapper_exception
// origin: languages/kotlin/tests/kotlin/test_result_patterns.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Result.success(0).mapCatching { throw Exception("mapped") }
            __check((value.isSuccess).toString(), "false")
            __check((value.exceptionOrNull()?.message).toString(), "mapped")
        }
