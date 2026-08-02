// vybe-test: kotlin/kotlin_set_construction/test_set_additional_distinct_elements
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_construction.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = mutableSetOf("a")
            s.add("b")
            s.add("a")
            __check((s.size).toString(), "2")
            __check((s.contains("b")).toString(), "true")
        }
