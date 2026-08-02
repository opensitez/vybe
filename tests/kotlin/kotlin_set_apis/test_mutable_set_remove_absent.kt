// vybe-test: kotlin/kotlin_set_apis/test_mutable_set_remove_absent
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = mutableSetOf(1, 2)
            val removed = set.remove(7)
            __check((removed).toString(), "false")
            __check((set.size).toString(), "2")
        }
