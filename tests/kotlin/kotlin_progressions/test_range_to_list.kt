// vybe-test: kotlin/kotlin_progressions/test_range_to_list
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = (1..3).toList()
            val a = (5 downTo 3).toList()
            __check((r.joinToString(",")).toString(), "1,2,3")
            __check((a.joinToString(",")).toString(), "5,4,3")
        }
