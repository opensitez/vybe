// vybe-test: kotlin/kotlin_associate_apis/test_associate_by_length
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val words = listOf("aa", "bb", "c", "ddd")
            val map = words.associateBy { it.length }
            __check((map[1]).toString(), "c")
            __check((map[3]).toString(), "ddd")
        }
