// vybe-test: kotlin/range_projection/test_char_range_projection
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 'a'..'e'
            __check((r.count()).toString(), "5")
            __check((r.first()).toString(), "a")
            __check((r.last()).toString(), "e")
            __check((r.contains('c')).toString(), "true")
            __check((('d' in r).toString()).toString(), "true")
        }
