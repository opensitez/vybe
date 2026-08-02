// vybe-test: kotlin/kotlin_pairs_apis/test_triple_not_equal
// origin: languages/kotlin/tests/kotlin/test_kotlin_pairs_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t1 = Triple(1, 2, 3)
            val t2 = Triple(3, 2, 1)
            __check((t1 == t2).toString(), "false")
        }
