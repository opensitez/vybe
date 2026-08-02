// vybe-test: kotlin/kotlin_set_apis/test_set_distinct_preserved_after_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf(1, 2, 3)
            val doubled = set.map { it * 2 }.toSet()
            __check((doubled.joinToString(",")).toString(), "2,4,6")
        }
