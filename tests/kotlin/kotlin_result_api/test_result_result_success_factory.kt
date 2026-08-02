// vybe-test: kotlin/kotlin_result_api/test_result_result_success_factory
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Result<String> = Result.success("ok")
            __check((value.isSuccess).toString(), "true")
            __check((value.getOrNull()).toString(), "ok")
        }
