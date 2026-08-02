// vybe-test: kotlin/constructor_chaining/test_constructor_list_param
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class WithList(val items: List<Int>)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val w = WithList(listOf(1, 2))
            __check((w.items.joinToString("|")).toString(), "1|2")
        }
