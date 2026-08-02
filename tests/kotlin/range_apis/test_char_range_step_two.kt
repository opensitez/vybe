// vybe-test: kotlin/range_apis/test_char_range_step_two
// origin: languages/kotlin/tests/kotlin/test_range_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = ('a'..'f').step(2)
            __check((r.toList().joinToString(",")).toString(), "a,c,e")
        }
