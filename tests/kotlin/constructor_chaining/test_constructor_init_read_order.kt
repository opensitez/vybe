// vybe-test: kotlin/constructor_chaining/test_constructor_init_read_order
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Bag(val base: Int) {
            val a: Int
            init { a = base + 1 }
            val b = a + 2
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Bag(3)
            __check((b.a).toString(), "4")
            __check((b.b).toString(), "6")
        }
