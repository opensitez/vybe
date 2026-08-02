// vybe-test: kotlin/collections_set/test_set_retain_all_with_self_is_noop
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3)
            __check((values.retainAll(values.toSet())).toString(), "false")
            __check((values.size).toString(), "3")
            __check((values.contains(1)).toString(), "true")
            __check((values.contains(3)).toString(), "true")
        }
