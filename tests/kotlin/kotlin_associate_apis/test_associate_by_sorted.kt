// vybe-test: kotlin/kotlin_associate_apis/test_associate_by_sorted
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = listOf("dog", "ant", "cat", "zoo").associateBy { it.length }
            val keys = map.keys.toList().sorted()
            __check((keys.joinToString(",")).toString(), "3,4")
        }
