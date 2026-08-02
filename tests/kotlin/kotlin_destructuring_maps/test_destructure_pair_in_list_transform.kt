// vybe-test: kotlin/kotlin_destructuring_maps/test_destructure_pair_in_list_transform
// origin: languages/kotlin/tests/kotlin/test_kotlin_destructuring_maps.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val entries = listOf(Pair("a", 3), Pair("b", 4))
            val keys = entries.map { (k, v) -> "${'$'}k${'$'}v" }
            __check((keys.joinToString("|")).toString(), "a3|b4")
        }
