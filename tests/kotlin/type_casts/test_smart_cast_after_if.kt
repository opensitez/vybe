// vybe-test: kotlin/type_casts/test_smart_cast_after_if
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item: Any = 55
            if (item is Int) {
                val value = item
                __check((value * 2).toString(), "110")
            }
        }
