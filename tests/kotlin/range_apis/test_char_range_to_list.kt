// vybe-test: kotlin/range_apis/test_char_range_to_list
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 'x'..'z'
            __check((r.toList().joinToString(",")).toString(), "x,y,z")
        }
