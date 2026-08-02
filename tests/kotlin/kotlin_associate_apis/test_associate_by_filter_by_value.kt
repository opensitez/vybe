// vybe-test: kotlin/kotlin_associate_apis/test_associate_by_filter_by_value
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf("x", "yy", "zzz").associateBy { it.length }.filterValues { it.startsWith("z") }
            __check((map.size).toString(), "1")
            __check((map[3]).toString(), "zzz")
        }
