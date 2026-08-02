// vybe-test: kotlin/kotlin_list_take_drop/test_drop_more
// origin: languages/kotlin/tests/kotlin/test_kotlin_list_take_drop.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = listOf("a", "b")
            __check((a.drop(5).toString()).toString(), "[]")
        }
