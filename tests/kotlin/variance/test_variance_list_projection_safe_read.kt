// vybe-test: kotlin/variance/test_variance_list_projection_safe_read
// origin: languages/kotlin/tests/kotlin/test_variance.rs

fun readFirst(items: List<out Number>): Int {
            return items.first().toInt()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((readFirst(listOf<Int>(5, 7))).toString(), "5")
            __check((readFirst(listOf<Long>(9L, 10L))).toString(), "9")
        }
