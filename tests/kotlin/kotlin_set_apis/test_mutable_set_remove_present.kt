// vybe-test: kotlin/kotlin_set_apis/test_mutable_set_remove_present
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = mutableSetOf(1, 2, 3)
            val removed = set.remove(2)
            __check((removed).toString(), "true")
            __check((set.size).toString(), "2")
            __check((set.contains(2)).toString(), "false")
        }
