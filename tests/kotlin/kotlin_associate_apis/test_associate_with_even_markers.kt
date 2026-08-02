// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_even_markers
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf(1, 2, 3, 4).associateWith { if (it % 2 == 0) "E" else "O" }
            __check((map[1]).toString(), "O")
            __check((map[2]).toString(), "E")
            __check((map[4]).toString(), "E")
        }
