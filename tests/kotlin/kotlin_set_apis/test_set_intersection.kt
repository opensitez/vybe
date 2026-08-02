// vybe-test: kotlin/kotlin_set_apis/test_set_intersection
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = setOf(1, 2, 3, 4)
            val b = setOf(3, 4, 5)
            val c = a.intersect(b)
            __check((c.size).toString(), "2")
            __check((c.joinToString(",")).toString(), "3,4")
        }
