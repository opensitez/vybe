// vybe-test: kotlin/kotlin_associate_apis/test_associate_with_count_chars
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val input = listOf("a", "bc", "de", "f")
            val map = input.associateWith { it.count() }
            val total = map.values.sum()
            __check((map.size).toString(), "4")
            __check((total).toString(), "7")
        }
