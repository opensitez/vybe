// vybe-test: kotlin/kotlin_associate_apis/test_associate_by_collision_last_wins
// origin: languages/kotlin/tests/kotlin/test_kotlin_associate_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val items = listOf("k1" to 1, "k1" to 2, "k2" to 3)
            val map = items.associateBy { it.first }
            __check((map.size).toString(), "2")
            __check((map["k1"]).toString(), "2")
        }
