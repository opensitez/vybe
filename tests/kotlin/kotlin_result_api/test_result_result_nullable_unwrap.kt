// vybe-test: kotlin/kotlin_result_api/test_result_result_nullable_unwrap
// origin: languages/kotlin/tests/kotlin/test_kotlin_result_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = runCatching { null as String? }
            val payload = value.getOrNull()
            __check((payload == null).toString(), "true")
        }
