// vybe-test: kotlin/kotlin_associate_apis/test_associate_pairs_to_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf("a" to 1, "b" to 2, "c" to 3).associateBy { it.first }
            __check((map.size).toString(), "3")
            __check((map["b"]).toString(), "2")
        }
