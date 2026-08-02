// vybe-test: kotlin/kotlin_set_apis/test_set_remove_all
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = mutableSetOf(1, 2, 3, 4)
            val changed = set.removeAll(listOf(1, 3))
            __check((changed).toString(), "true")
            __check((set.size).toString(), "2")
            __check((set.contains(3)).toString(), "false")
        }
