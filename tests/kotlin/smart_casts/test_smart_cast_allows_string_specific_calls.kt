// vybe-test: kotlin/smart_casts/test_smart_cast_allows_string_specific_calls
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = "Rust"
            val upper = if (value is String) {
                value.uppercase()
            } else {
                "none"
            }
            __check((upper).toString(), "RUST")
        }
