// vybe-test: kotlin/smart_casts/test_as_cast_success
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = "kotlin"
            val text = value as String
            __check((text.length).toString(), "6")
        }
