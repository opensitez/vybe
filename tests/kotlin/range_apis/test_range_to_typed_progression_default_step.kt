// vybe-test: kotlin/range_apis/test_range_to_typed_progression_default_step
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 3 until 3
            __check((r.isEmpty()).toString(), "true")
            __check((r.toList().size).toString(), "0")
        }
