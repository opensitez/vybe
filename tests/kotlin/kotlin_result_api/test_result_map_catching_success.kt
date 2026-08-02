// vybe-test: kotlin/kotlin_result_api/test_result_map_catching_success
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ok = runCatching { 5 }
                .mapCatching { if (it == 5) "five" else throw Exception("x") }
            __check((ok.getOrNull()).toString(), "five")
            __check((ok.exceptionOrNull() == null).toString(), "true")
        }
