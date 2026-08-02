// vybe-test: kotlin/kotlin_result_api/test_result_map_success_only
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = runCatching { 4 }
                .map { it * 10 }
            val bad = runCatching<Int> { throw Exception("x") }
                .map { it * 2 }
            __check((ok.getOrNull()).toString(), "40")
            __check((bad.getOrNull() == null).toString(), "true")
        }
