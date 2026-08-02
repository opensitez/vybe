// vybe-test: kotlin/kotlin_set_apis/test_set_retain_all
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = mutableSetOf(1, 2, 3, 4)
            val changed = set.retainAll(listOf(2, 4))
            __check((changed).toString(), "true")
            __check((set.joinToString(",")).toString(), "2,4")
        }
