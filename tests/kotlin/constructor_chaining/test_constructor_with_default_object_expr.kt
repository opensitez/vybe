// vybe-test: kotlin/constructor_chaining/test_constructor_with_default_object_expr
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Holder(val value: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h: Holder = Holder(3)
            __check((h.value).toString(), "3")
        }
